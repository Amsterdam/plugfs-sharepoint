from collections.abc import AsyncIterator

import httpx
import pytest
from plugfs.filesystem import Directory, Filesystem

from plugfs_sharepoint import SharepointAdapter, SharepointFile
from plugfs_sharepoint.adapter import _is_retryable_http_error


class TokenFactory:
    async def __call__(self) -> str:
        return "test-token"


@pytest.fixture
def transport() -> httpx.MockTransport:
    calls: list[tuple[str, str]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append((request.method, str(request.url.path)))

        if request.url.path == "/v1.0/sites/example.sharepoint.com:/sites/MySite":
            return httpx.Response(200, json={"id": "site-1"})

        if request.url.path == "/v1.0/sites/site-1/drive":
            return httpx.Response(200, json={"id": "drive-1"})

        if request.url.path == "/v1.0/drives/drive-1/root":
            return httpx.Response(200, json={"id": "root-id", "size": 0})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs":
            return httpx.Response(200, json={"id": "folder-id", "size": 0})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs:/content":
            return httpx.Response(200, content=b"folder-bytes")

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt":
            return httpx.Response(200, json={"id": "file-id", "size": 12})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt:/content":
            return httpx.Response(200, content=b"hello world")

        if request.url.path == "/v1.0/drives/drive-1/items/folder-id/children":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {"name": "report.txt", "size": 12},
                        {"name": "docs", "folder": {}},
                    ]
                },
            )

        if request.url.path == "/v1.0/drives/drive-1/items/file-id/children":
            return httpx.Response(200, json={"value": []})

        if request.url.path == "/v1.0/drives/drive-1/items/root-id/children":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {"name": "report.txt", "size": 12},
                        {"name": "docs", "folder": {}},
                    ]
                },
            )

        return httpx.Response(404, json={"error": "not found"})

    transport.calls = calls  # ty: ignore[unresolved-attribute]
    return httpx.MockTransport(handler)


@pytest.fixture
def client(transport: httpx.MockTransport) -> httpx.AsyncClient:
    return httpx.AsyncClient(
        base_url="https://graph.microsoft.com/v1.0",
        transport=transport,
    )


def retry_transport() -> httpx.MockTransport:
    state = {"site_attempts": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/sites/example.sharepoint.com:/sites/MySite":
            state["site_attempts"] += 1
            if state["site_attempts"] == 1:
                return httpx.Response(429, json={"error": "retry"})
            return httpx.Response(200, json={"id": "site-1"})

        if request.url.path == "/v1.0/sites/site-1/drive":
            return httpx.Response(200, json={"id": "drive-1"})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt":
            return httpx.Response(200, json={"id": "file-id", "size": 12})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt:/content":
            return httpx.Response(200, content=b"hello world")

        return httpx.Response(404, json={"error": "not found"})

    return httpx.MockTransport(handler)


@pytest.mark.anyio
async def test_create_and_read_file(client: httpx.AsyncClient) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    file = await adapter.get_file("/docs/report.txt")
    assert isinstance(file, SharepointFile)
    assert await file.read() == b"hello world"
    assert await file.size == 12
    assert await _collect_bytes(await file.get_iterator()) == b"hello world"

    await adapter.aclose()


@pytest.mark.anyio
async def test_list_maps_directories_and_files(client: httpx.AsyncClient) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    items = await adapter.list("/docs")

    assert len(items) == 2
    assert isinstance(items[0], SharepointFile)
    assert items[0].path == "/docs/report.txt"
    assert isinstance(items[1], Directory)
    assert items[1].path == "/docs/docs"

    await adapter.aclose()


@pytest.mark.anyio
async def test_create_uses_graph_resolution(monkeypatch: pytest.MonkeyPatch) -> None:
    mock_client = httpx.AsyncClient(
        base_url="https://graph.microsoft.com/v1.0",
        transport=transport_for_create(),
    )

    monkeypatch.setattr(
        "plugfs_sharepoint.adapter.httpx.AsyncClient",
        lambda **_: mock_client,
    )

    adapter = await SharepointAdapter.create(
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )

    file = await adapter.get_file("/docs/report.txt")
    assert isinstance(file, SharepointFile)
    assert await file.read() == b"hello world"

    await adapter.aclose()


@pytest.mark.anyio
async def test_get_iterator_streams_and_closes_response(
    client: httpx.AsyncClient,
) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    iterator = await adapter.get_iterator("/docs/report.txt")
    assert await _collect_bytes(iterator) == b"hello world"

    await adapter.aclose()


@pytest.mark.anyio
async def test_create_closes_client_when_initialization_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FailingClient:
        def __init__(self) -> None:
            self.closed = False

        async def aclose(self) -> None:
            self.closed = True

    failing_client = FailingClient()

    monkeypatch.setattr(
        "plugfs_sharepoint.adapter.httpx.AsyncClient",
        lambda **_: failing_client,
    )

    async def raise_init(self: SharepointAdapter) -> None:
        raise RuntimeError("boom")

    monkeypatch.setattr(SharepointAdapter, "_initialize", raise_init)

    with pytest.raises(RuntimeError):
        await SharepointAdapter.create(
            TokenFactory(),
            "example.sharepoint.com",
            "/sites/MySite",
        )

    assert failing_client.closed is True


@pytest.mark.anyio
async def test_retry_on_429(client: httpx.AsyncClient) -> None:
    retry_client = httpx.AsyncClient(
        base_url="https://graph.microsoft.com/v1.0",
        transport=retry_transport(),
    )
    adapter = SharepointAdapter(
        retry_client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )

    await adapter._initialize()

    file = await adapter.get_file("/docs/report.txt")
    assert await file.read() == b"hello world"

    await adapter.aclose()


@pytest.mark.anyio
async def test_read_only_methods_raise(client: httpx.AsyncClient) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    with pytest.raises(NotImplementedError):
        await adapter.write("/docs/report.txt", b"x")

    with pytest.raises(NotImplementedError):
        await adapter.write_iterator("/docs/report.txt", _empty_iterator())

    with pytest.raises(NotImplementedError):
        await adapter.makedirs("/docs/new")

    with pytest.raises(NotImplementedError):
        await adapter.delete("/docs/report.txt")

    with pytest.raises(NotImplementedError):
        await (await adapter.get_file("/docs/report.txt")).delete()

    await adapter.aclose()


@pytest.mark.anyio
async def test_filesystem_wrapper_uses_adapter(client: httpx.AsyncClient) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    filesystem = Filesystem(adapter)
    file = await filesystem.get_file("/docs/report.txt")

    assert await file.read() == b"hello world"

    await adapter.aclose()


@pytest.mark.anyio
async def test_root_path_and_helpers(client: httpx.AsyncClient) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    await adapter._initialize()

    root = await adapter.get_file("/")
    assert root.path == "/"

    items = await adapter.list("/")
    assert len(items) == 2

    assert adapter._normalize_path("/") == "/"
    assert adapter._normalize_path("docs") == "/docs"
    assert adapter._to_path("/docs/", "name") == "/docs/name"

    await adapter.aclose()


@pytest.mark.anyio
async def test_get_drive_item_requires_initialization(
    client: httpx.AsyncClient,
) -> None:
    adapter = SharepointAdapter(
        client,
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )

    with pytest.raises(RuntimeError):
        await adapter._get_drive_item("/docs/report.txt")


def test_retry_helper_rejects_non_http_errors() -> None:
    assert _is_retryable_http_error(ValueError("boom")) is False


def _empty_iterator() -> AsyncIterator[bytes]:
    async def iterator() -> AsyncIterator[bytes]:
        if False:
            yield b""

    return iterator()


def _single_chunk_iterator(chunk: bytes) -> AsyncIterator[bytes]:
    async def iterator() -> AsyncIterator[bytes]:
        yield chunk

    return iterator()


async def _collect_bytes(iterator: AsyncIterator[bytes]) -> bytes:
    chunks: list[bytes] = []
    async for chunk in iterator:
        chunks.append(chunk)
    return b"".join(chunks)


def transport_for_create() -> httpx.MockTransport:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/sites/example.sharepoint.com:/sites/MySite":
            return httpx.Response(200, json={"id": "site-1"})

        if request.url.path == "/v1.0/sites/site-1/drive":
            return httpx.Response(200, json={"id": "drive-1"})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt":
            return httpx.Response(200, json={"id": "file-id", "size": 12})

        if request.url.path == "/v1.0/drives/drive-1/root:/docs/report.txt:/content":
            return httpx.Response(200, content=b"hello world")

        if request.url.path == "/v1.0/drives/drive-1/root":
            return httpx.Response(200, json={"id": "root-id", "size": 0})

        if request.url.path == "/v1.0/drives/drive-1/items/root-id/children":
            return httpx.Response(200, json={"value": []})

        return httpx.Response(404, json={"error": "not found"})

    return httpx.MockTransport(handler)

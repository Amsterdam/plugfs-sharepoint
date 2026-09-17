from __future__ import annotations

from collections.abc import AsyncIterator, Sequence
from dataclasses import dataclass
from typing import Any, Protocol, cast, final, runtime_checkable

import httpx
from plugfs.filesystem import (
    Adapter,
    Directory,
    DirectoryListing,
    File,
    _FilesystemItem,
)
from tenacity import (
    retry,
    retry_if_exception,
    stop_after_attempt,
    wait_exponential,
)


@runtime_checkable
class AccessTokenFactory(Protocol):
    async def __call__(self) -> str: ...


def _is_retryable_http_error(error: BaseException) -> bool:
    if not isinstance(error, httpx.HTTPStatusError):
        return False

    status_code = cast(httpx.Response, cast(Any, error).response).status_code
    return status_code == 429 or 500 <= status_code < 600


@dataclass(slots=True)
class _SharepointSite:
    site_id: str
    drive_id: str


@final
class SharepointFile(File):
    _adapter: SharepointAdapter

    def __init__(self, path: str, adapter: SharepointAdapter) -> None:
        super().__init__(path)
        self._adapter = adapter

    @property
    async def size(self) -> int:
        return await self._adapter.get_size(self._path)

    async def read(self) -> bytes:
        return await self._adapter.read(self._path)

    async def get_iterator(self) -> AsyncIterator[bytes]:
        return await self._adapter.get_iterator(self._path)

    async def delete(self) -> None:
        raise NotImplementedError("SharePoint adapter is read-only")


@final
class SharepointAdapter(Adapter):
    _client: httpx.AsyncClient
    _token_factory: AccessTokenFactory
    _site_host: str
    _site_path: str
    _site: _SharepointSite | None

    def __init__(
        self,
        client: httpx.AsyncClient,
        token_factory: AccessTokenFactory,
        site_host: str,
        site_path: str,
    ) -> None:
        self._client = client
        self._token_factory = token_factory
        self._site_host = site_host
        self._site_path = site_path
        self._site = None

    @classmethod
    async def create(
        cls,
        token_factory: AccessTokenFactory,
        site_host: str,
        site_path: str,
    ) -> SharepointAdapter:
        client = httpx.AsyncClient(base_url="https://graph.microsoft.com/v1.0")
        adapter = cls(client, token_factory, site_host, site_path)

        try:
            await adapter._initialize()
        except Exception:
            await client.aclose()
            raise

        return adapter

    async def _initialize(self) -> None:
        site_response = await self._get(f"/sites/{self._site_host}:{self._site_path}")
        site_data = cast(dict[str, Any], site_response.json())
        site_id = cast(str, site_data["id"])

        drive_response = await self._get(f"/sites/{site_id}/drive")
        drive_data = cast(dict[str, Any], drive_response.json())

        self._site = _SharepointSite(
            site_id=site_id,
            drive_id=cast(str, drive_data["id"]),
        )

    async def aclose(self) -> None:
        await self._client.aclose()

    async def list(self, path: str) -> DirectoryListing:
        drive_item = await self._get_drive_item(path)
        children = await self._get_drive_children(cast(str, drive_item["id"]))

        items: list[_FilesystemItem] = []
        for child in children:
            child_name = cast(str, child["name"])
            if "folder" in child:
                items.append(Directory(self._to_path(path, child_name)))
            else:
                items.append(SharepointFile(self._to_path(path, child_name), self))

        return items

    async def read(self, path: str) -> bytes:
        response = await self._get_drive_item_content(path)
        try:
            return await response.aread()
        finally:
            await response.aclose()

    async def get_iterator(self, path: str) -> AsyncIterator[bytes]:
        response = await self._get_drive_item_content(path)

        async def iterate() -> AsyncIterator[bytes]:
            try:
                async for chunk in response.aiter_bytes():
                    yield chunk
            finally:
                await response.aclose()

        return iterate()

    async def get_file(self, path: str) -> SharepointFile:
        await self._get_drive_item(path)
        return SharepointFile(path, self)

    async def write(self, path: str, data: bytes) -> File:
        raise NotImplementedError("SharePoint adapter is read-only")

    async def write_iterator(self, path: str, iterator: AsyncIterator[bytes]) -> File:
        raise NotImplementedError("SharePoint adapter is read-only")

    async def makedirs(self, path: str) -> None:
        raise NotImplementedError("SharePoint adapter is read-only")

    async def delete(self, path: str) -> None:
        raise NotImplementedError("SharePoint adapter is read-only")

    @retry(
        retry=retry_if_exception(_is_retryable_http_error),
        wait=wait_exponential(multiplier=1, min=1, max=30),
        stop=stop_after_attempt(5),
        reraise=True,
    )
    async def _get(self, path: str) -> httpx.Response:
        token = await self._token_factory()
        response = await self._client.get(
            path,
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        return response

    async def get_size(self, path: str) -> int:
        item = await self._get_drive_item(path)
        return int(cast(int | str, item["size"]))

    async def _get_drive_item(self, path: str) -> dict[str, Any]:
        site = self._require_site()
        normalized_path = self._normalize_path(path)
        if normalized_path == "/":
            response = await self._get(f"/drives/{site.drive_id}/root")
            return cast(dict[str, Any], response.json())

        response = await self._get(f"/drives/{site.drive_id}/root:{normalized_path}")
        return cast(dict[str, Any], response.json())

    async def _get_drive_item_content(self, path: str) -> httpx.Response:
        site = self._require_site()
        normalized_path = self._normalize_path(path)
        return await self._get(
            f"/drives/{site.drive_id}/root:{normalized_path}:/content"
        )

    async def _get_drive_children(self, item_id: str) -> Sequence[dict[str, Any]]:
        site = self._require_site()
        response = await self._get(f"/drives/{site.drive_id}/items/{item_id}/children")
        payload = cast(dict[str, Any], response.json())
        return cast(Sequence[dict[str, Any]], payload.get("value", []))

    def _require_site(self) -> _SharepointSite:
        if self._site is None:
            raise RuntimeError("SharePoint adapter not initialized")
        return self._site

    @staticmethod
    def _normalize_path(path: str) -> str:
        if path == "/":
            return path
        if not path.startswith("/"):
            return f"/{path}"
        return path

    @staticmethod
    def _to_path(parent: str, name: str) -> str:
        if parent.endswith("/"):
            return f"{parent}{name}"
        return f"{parent}/{name}"

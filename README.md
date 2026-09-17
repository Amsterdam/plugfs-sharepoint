# plugfs-sharepoint

SharePoint adapter for [`plugfs`](https://github.com/Amsterdam/plugfs).

## Install

```shell
uv add plugfs-sharepoint
```

or:

```shell
pip install plugfs-sharepoint
```

## What it does

This package provides a SharePoint adapter for `plugfs` using Microsoft Graph.

## Current status

The filesystem is currently read-only.

Support for `write`, `write_iterator`, `makedirs`, and `delete` is not implemented yet and will be added when needed.

## Usage

```python
from plugfs.filesystem import Filesystem
from plugfs_sharepoint import SharepointAdapter


class TokenFactory:
    async def __call__(self) -> str:
        return "access-token"


async def build_filesystem() -> Filesystem:
    adapter = await SharepointAdapter.create(
        TokenFactory(),
        "example.sharepoint.com",
        "/sites/MySite",
    )
    return Filesystem(adapter)
```

Quick smoke test:

```python
async def smoke_test(filesystem: Filesystem) -> None:
    items = await filesystem.list("/")
    for item in items:
        print(item.path)
```

## Contract

- `site_host` must be host only, for example `example.sharepoint.com`
- `site_path` must be site path only, for example `/sites/MySite`
- token acquisition stays with caller
- filesystem is currently read-only
- `write`, `write_iterator`, `makedirs`, and `delete` are not implemented yet

## Notes

The adapter uses Microsoft Graph to resolve the site and default document library drive.

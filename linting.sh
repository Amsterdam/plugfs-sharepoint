#!/usr/bin/env bash

set -e
set -u
set -x

uvx ruff check src tests
uvx ruff format --check src tests
uvx ty check .

#!/usr/bin/env bash

set -u
set -x

uvx ruff check --fix src tests
uvx ruff format src tests

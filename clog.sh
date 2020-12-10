#!/bin/bash
(git tag -l |rg '^v\d+') \
  && clog || clog --setversion v$(toml get Cargo.toml package.version |sed 's/"//g') \
  && git add CHANGELOG.md && git commit -m "chore: add changelog for pre release"

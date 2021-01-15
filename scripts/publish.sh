#!/bin/bash
set -e -u

version=$(cat package.json | jq -r '.version')
git tag v$version -f

yarn build

publish_opts=$(echo $version | grep -q beta && echo "--tag beta" || true)
yarn publish $publish_opts --new-version $version

git push
git push --tags

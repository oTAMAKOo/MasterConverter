

# mac, linux のビルド確認（docker）

https://hub.docker.com/_/mono/

docker の mono:5.20.1.19 コンテナでビルドができます。

```
$ docker run -it --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 bash
/src# mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll

もしくは

$ docker run --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll
```

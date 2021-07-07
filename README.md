
# コンバートツールの利用方法

WIP

## コンバートツールの引数挙動

```
--mode import：ClassSchema.xlsxをコピー後リネームして、レコードファイルを流し込んで編集用xlsxを生成
--mode export：importで作成された編集用xlsxからレコード情報を生成
--mode build：ClassSchema.xlsxとレコード情報からクライアント配信用(.master)とサーバー読み込み用(.yml)を生成
```

# mac, linux のビルド確認（docker）

https://hub.docker.com/_/mono/

docker の mono:5.20.1.19 コンテナでビルドができます。

```
$ docker run -it --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 bash
/src# mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll

もしくは

$ docker run --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll
```

# テストフィクスチャ

| ファイル | 出所 |
| --- | --- |
| `sample.md` | `sample/sample.md` のコピー。全 Markdown 要素を含む |
| `template.pptx` | `sample/template.pptx` のコピー。16:9・11 レイアウト |
| `config.yaml` | `sample/config.yaml` のコピー |

`sample/` は `.gitignore` で除外されているため、テストが単体で完結するように
ここに複製しています。上流の `sample/` を更新したら、あわせてコピーし直してください。

```bash
cp sample/sample.md sample/template.pptx sample/config.yaml tests/fixtures/
```

`.difyignore` が `tests/` を除外するので、これらは `.difypkg` には含まれません。

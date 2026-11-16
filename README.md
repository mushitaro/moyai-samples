# moyai-samples

**All data in this repository is fictional (【架空】).** The facilities, municipalities, companies, people, equipment, model numbers, serial numbers and amounts were invented for testing. They do not describe any real facility or organization.

**このリポジトリの中身はすべて架空（【架空】）です。** 施設・自治体・会社・人・機器・型式・製造番号・金額はすべて試験のために作ったもので、実在の施設や団体を表すものではありません。

---

## English

Test samples for evaluating the reader of [MOYAI](https://github.com/mushitaro/moyai), a maintenance-planning tool. Each sample is the maintenance paperwork of one fictional municipal waste incineration plant, written the way a site would keep it (Excel, CSV system exports, text PDFs and scanned handwritten pages), together with the answer key that the evaluation harness reads.

### Facilities

| Folder | Facility (fictional) | Furnace type | Capacity | Fiscal years | Documents |
|---|---|---|---|---|---|
| `waste-incineration/f01-asagi` | Asagi City Incineration Plant | Stoker, continuous | 110 t × 2 | 1997–2026 | 908 |
| `waste-incineration/f02-hibarino` | Hibarino Clean Center | Stoker, continuous | 105 t × 2 | 2016–2026 | 105 |
| `waste-incineration/f03-kunugidai` | Kunugidai Environment Center | Stoker, continuous | 100 t × 3 | 2001–2026 | 375 |
| `waste-incineration/f04-shiosai` | Shiosai Town Waste Plant | Stoker, semi-continuous | 20 t/16 h × 2 | 2012–2026 | 204 |
| `waste-incineration/f05-yamabuki` | Yamabuki Regional Union, Plant No. 2 | Stoker, semi-continuous | 30 t/16 h × 2 | 2021–2025 | 32 |
| `waste-incineration/f06-minato` | Minato Waterfront Plant | Stoker, continuous | 200 t × 3 | 2004–2026 | 267 |
| `waste-incineration/f07-kazahaya` | Kazahaya Eco Plaza | Stoker, continuous | 150 t × 2 | 2015–2026 | 151 |
| `waste-incineration/f08-izumigawa` | Izumigawa Cleaning Center | Fluidized bed (shut down FY2024) | 75 t × 2 | 2005–2024 | 110 |
| `waste-incineration/f09-tokiwa` | Tokiwa North Clean Center | Fluidized-bed gasification & melting | 95 t × 2 | 2009–2026 | 158 |
| `waste-incineration/f10-akanedai` | Akanedai Resource Center | Shaft-furnace gasification & melting | 120 t × 2 | 2011–2026 | 193 |
| `waste-incineration/f11-sakurazaka` | Sakurazaka Incinerator | Stoker, continuous | 60 t × 2 | 2019–2026 | 62 |

Total: 2,565 documents. Each facility has six kinds of documents: long-term maintenance plan, construction schedule, daily work report, completion report, purchase ledger and equipment ledger. Every facility has at least 250 managed pieces of equipment.

### Layout of a facility folder

```
<facility>/
  files/          the documents (flat, no subfolders; names start with 【架空】)
  user.md         what the facility's maintenance staff would answer if asked
  README.md       deliberate variations and inconsistencies in this facility's documents
  answer/
    world.json        answer key, schema "moyai.sample-world/0"
    world.extra.json  facts kept from the earlier answer format that the harness does not read
    files/<stem>.json per-document answer (records with sheet, r, kind, fields)
  gen/            the generator and check.py
    _answer_v1/       the earlier-format answers that to_v2.py converts
```

### Answer format (`moyai.sample-world/0`)

`answer/world.json` holds `facility`, `profile`, `maker`, `assets[]`, `works[]`, `links[]` (which row of which document records which work), `errata[]` (only `omitMark` and `lastYearUnreflected`) and `files[]`. `answer/files/<stem>.json` lists the data rows of one document. Plan (星取表) rows carry `fields.path`, `fields.marks` and `fields.plans`. Other documents carry `fields.manDays` (person-days, 1 person-day = 8 hours) and `fields.days`. Photo and scanned pages keep `seq` and `image` and are skipped by the harness.

### Rebuilding

Each facility's `gen/` folder rebuilds that facility. To regenerate the answer files from `gen/_answer_v1/`, run `python3 to_v2.py` in `gen/`. To check the answers against the documents and confirm the 【架空】 marks, run `python3 check.py`. The full document generator is also in `gen/` (`run_all.py`, `run_full.py` or `build.py`, depending on the facility). It needs Python 3, openpyxl, python-docx or reportlab, and LibreOffice. Rebuilding the documents can change PDF timestamps, so do it only when you intend to replace `files/`.

## 日本語

保全計画のツール [MOYAI](https://github.com/mushitaro/moyai) の読み手を評価するための試験用サンプルです。架空の一般廃棄物焼却施設 1 か所ずつの保全の資料を、現場が持っている形で作りました（Excel、システムが出力した CSV、文字のある PDF、手書きをスキャンした頁）。物差しが読む正解も付けています。

### 施設の一覧

| フォルダ | 施設名（架空） | 炉形式 | 能力 | 対象年度 | 資料の数 |
|---|---|---|---|---|---|
| `waste-incineration/f01-asagi` | あさぎ市清掃工場 | ストーカ・全連続 | 110t×2 | 1997–2026 | 908 |
| `waste-incineration/f02-hibarino` | ひばり野クリーンセンター | ストーカ・全連続 | 105t×2 | 2016–2026 | 105 |
| `waste-incineration/f03-kunugidai` | くぬぎ台環境センター | ストーカ・全連続 | 100t×3 | 2001–2026 | 375 |
| `waste-incineration/f04-shiosai` | しおさい町ごみ処理場 | ストーカ・准連続 | 20t/16h×2 | 2012–2026 | 204 |
| `waste-incineration/f05-yamabuki` | やまぶき広域処理組合 第二工場 | ストーカ・准連続 | 30t/16h×2 | 2021–2025 | 32 |
| `waste-incineration/f06-minato` | みなと臨海工場 | ストーカ・全連続 | 200t×3 | 2004–2026 | 267 |
| `waste-incineration/f07-kazahaya` | かざはやエコプラザ | ストーカ・全連続 | 150t×2 | 2015–2026 | 151 |
| `waste-incineration/f08-izumigawa` | いずみ川清掃センター | 流動床（2024 年度で停止） | 75t×2 | 2005–2024 | 110 |
| `waste-incineration/f09-tokiwa` | ときわ北部クリーンセンター | 流動床式ガス化溶融 | 95t×2 | 2009–2026 | 158 |
| `waste-incineration/f10-akanedai` | あかね台資源化センター | シャフト炉式ガス化溶融 | 120t×2 | 2011–2026 | 193 |
| `waste-incineration/f11-sakurazaka` | さくら坂焼却場 | ストーカ・全連続 | 60t×2 | 2019–2026 | 62 |

資料は合計 2,565 です。どの施設にも 6 種の資料（中長期整備計画・工程表・作業日報・工事報告書・発注台帳・機器台帳）があり、管理対象の機器は 250 台以上です。

### 施設フォルダの形

```
<施設>/
  files/          資料（サブフォルダなし。ファイル名は【架空】で始まる）
  user.md         施設の保全担当者が聞かれたら答えられること
  README.md       この施設の資料にわざと入れた揺れと食い違いの一覧
  answer/
    world.json        正解（schema "moyai.sample-world/0"）
    world.extra.json  前の形の正解にあって物差しが読まない事実
    files/<stem>.json 資料ごとの正解（記録は sheet・r・kind・fields）
  gen/            生成器と check.py
    _answer_v1/       to_v2.py が変換する前の形の正解
```

### 正解の形（`moyai.sample-world/0`）

`answer/world.json` には次の欄があります。物差しが読むのはこれだけです。

- `facility`・`profile`・`maker`
- `assets[]`
- `works[]`
- `links[]`：どの資料のどの行がどの作業か
- `errata[]`：`omitMark` と `lastYearUnreflected` だけ
- `files[]`

`answer/files/<stem>.json` は、資料 1 つのデータの行の一覧です。

- 星取表の行は `fields.path`・`fields.marks`・`fields.plans` を持ちます。
- ほかの資料の行は `fields.manDays`（人日。1 人日＝8 時間）と `fields.days` を持ちます。
- 写真とスキャンの頁は `seq` と `image` を持ち、物差しはこれを読まずに外します。

### 作り直し方

各施設の `gen/` で、その施設を作り直せます。

- 正解：`gen/` で `python3 to_v2.py` を走らせると、`gen/_answer_v1/` から作り直せます。
- 確かめ：`python3 check.py` で、正解の行が資料の中身と合うか、【架空】が入っているかを確かめます。
- 資料：生成器そのもの（施設によって `run_all.py`・`run_full.py`・`build.py`）も `gen/` にあります。Python 3、openpyxl、python-docx か reportlab、LibreOffice が要ります。資料を作り直すと PDF の作成日時などが変わるので、`files/` を入れ替えるときだけ走らせてください。

## License

[Apache License 2.0](LICENSE)

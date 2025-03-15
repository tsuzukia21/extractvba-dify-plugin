# VBA Extractor Dify Plugin

## 概要 / Overview

このプラグインは、ExcelマクロファイルからVBAコードを抽出するためのDifyプラグインです。  
This plugin is designed to extract VBA code from Excel macro files for the Dify platform.

マクロを含むExcelファイルをアップロードするだけで、内部のVBAコードをすべて抽出し、モジュール名とコードを整理して表示します。  
Simply upload Excel files containing macros, and the plugin will extract all internal VBA code, organizing and displaying module names and code.

## サポートしているファイル形式 / Supported File Formats

以下のファイル形式からVBAコードを抽出できます：  
The following file formats are supported for VBA code extraction:

* Excel マクロ有効ブック (.xlsm)  
  Excel Macro-Enabled Workbook (.xlsm)
* Excel マクロ有効テンプレート (.xltm)  
  Excel Macro-Enabled Template (.xltm)
* Excel アドイン (.xlam)  
  Excel Add-In (.xlam)
* 旧形式のExcelファイル (.xls)  
  Legacy Excel Files (.xls)

## Difyでの使用方法 / Usage in Dify

### パラメータ / Parameters

プラグインは以下のパラメータを受け付けます：  
The plugin accepts the following parameters:

```yaml
files:
  type: files
  required: true
  description: "マクロを含むExcelファイル（複数可）/ Excel files containing macros"
```

### 応答形式 / Response Format

プラグインは以下の形式でJSON応答を返します：  
The plugin returns a JSON response in the following format:

```json
{
  "status": "success",
  "results": [
    {
      "file_name": "example.xlsm",
      "found_macros": true,
      "modules": [
        {
          "module_name": "Module1",
          "code": "Sub Example()\n    MsgBox \"Hello World\"\nEnd Sub"
        },
        {
          "module_name": "Class1",
          "code": "Public Function TestFunction() As String\n    TestFunction = \"Test\"\nEnd Function"
        }
      ],
      "summary": "Found 2 VBA modules"
    }
  ]
}
```

エラーが発生した場合：  
In case of an error:

```json
{
  "status": "success",
  "results": [
    {
      "file_name": "example.xlsm",
      "found_macros": false,
      "modules": [],
      "error": "Error message",
      "summary": "Error while processing the file"
    }
  ]
}
```


## 特徴 / Features

* 複数ファイル処理：一度に複数のマクロファイルを処理できます  
  Batch Processing: Process multiple macro files in a single request
* モジュール分離：各VBAモジュールを名前付きで個別に抽出します  
  Module Separation: Extracts each VBA module individually with its name
* エラー処理：ファイル処理中のエラーを明確に報告します  
  Error Handling: Clearly reports errors encountered during file processing
* 自動クリーンアップ：一時ファイルは自動的に管理・削除されます  
  Automatic Cleanup: Temporary files are automatically managed and removed

## 技術的詳細 / Technical Details

* [oletools](https://github.com/decalage2/oletools)ライブラリを利用してVBAコードを抽出  
  Uses the [oletools](https://github.com/decalage2/oletools) library to extract VBA code

## 制限事項 / Limitations

* パスワード保護されたVBAプロジェクトからはコードを抽出できません  
  Cannot extract code from password-protected VBA projects
* 非常に大きなマクロファイルは処理時間が長くなる場合があります  
  Very large macro files may take longer to process
* コードのフォーマットは元のままで提供され、整形やハイライトは行いません  
  Code formatting is provided as-is, without any reformatting or highlighting

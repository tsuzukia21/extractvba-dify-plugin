from collections.abc import Generator
from typing import Any
import os
import tempfile
import base64
from oletools.olevba import VBA_Parser
import logging
import requests

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

class ExtractVBATool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        """
        Extract VBA code from Excel macro files
        """
        try:
            # filesパラメータをリストとして取得
            files = tool_parameters.get('files', [])
            
            # 空のリストの場合はエラーを返す
            if not files:
                yield self.create_text_message("No files provided")
                yield self.create_json_message({
                    "status": "error",
                    "message": "No files provided",
                    "results": []
                })
                return
            
            all_results = []
            json_results = []
            
            # 各ファイルを処理
            for file in files:
                try:
                    file_extension = '.xlsm'  # デフォルト拡張子
                    if hasattr(file, 'extension') and file.extension:
                        file_extension = file.extension
                    
                    # ファイルURLを取得して処理
                    file_url = None
                    if hasattr(file, 'url'):
                        file_url = file.url
                    elif isinstance(file, str) and file.startswith('http'):
                        file_url = file
                    
                    # URLからファイルをダウンロード
                    if file_url:
                        response = requests.get(file_url)
                        if response.status_code == 200:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as temp_file:
                                temp_file.write(response.content)
                                temp_file_path = temp_file.name
                        else:
                            raise ValueError(f"Failed to download file from URL: {response.status_code}")
                    
                    try:
                        # VBAコードの抽出
                        vba_result = self._extract_vba_code(temp_file_path)
                        vba_result["file_name"] = file.filename
                        all_results.append(vba_result)
                        
                        # JSON結果に追加
                        json_results.append({
                            "filename": file.filename,
                            "original_format": file_extension.lstrip('.'),
                            "found_macros": vba_result["found_macros"],
                            "summary": vba_result["summary"],
                            "modules": vba_result["modules"],
                            "status": "success"
                        })
                    finally:
                        # 一時ファイルの削除
                        if os.path.exists(temp_file_path):
                            os.unlink(temp_file_path)
                
                except Exception as e:
                    error_msg = f"Error processing file {file.filename}: {str(e)}"
                    yield self.create_text_message(error_msg)
                    json_results.append({
                        "filename": file.filename,
                        "original_format": file_extension.lstrip('.') if hasattr(file, 'extension') else "unknown",
                        "error": error_msg,
                        "status": "error"
                    })
            
            # 結果をJSONとして返す
            json_response = {
                "status": "success" if len(all_results) > 0 else "error",
                "total_files": len(files),
                "successful_extractions": len(all_results),
                "results": json_results
            }
            yield self.create_json_message(json_response)
            
            # テキスト結果も返す（後方互換性のため）
            if len(all_results) == 0:
                yield self.create_text_message("No VBA code could be extracted from the provided files.")
            elif len(all_results) == 1:
                yield self.create_text_message(f"Extracted VBA code from {all_results[0]['file_name']}: {all_results[0]['summary']}")
            else:
                combined_results = ""
                for idx, result in enumerate(all_results, 1):
                    combined_results += f"\n{'='*50}\n"
                    combined_results += f"ファイル {idx}: {result['file_name']}\n"
                    combined_results += f"{'='*50}\n\n"
                    combined_results += f"概要: {result['summary']}\n\n"
                    
                    if result['found_macros']:
                        for module in result['modules']:
                            combined_results += f"モジュール: {module['module_name']}\n"
                            combined_results += f"コード:\n{module['code']}\n\n"
                
                yield self.create_text_message(combined_results.strip())
            
        except Exception as e:
            error_message = f"Error extracting VBA code: {str(e)}"
            yield self.create_text_message(error_message)
    
    def _extract_vba_code(self, file_path: str) -> dict:
        """
        Use oletools to extract VBA code from the file
        """
        result = {}
        
        try:
            vba_parser = VBA_Parser(file_path)
            
            if vba_parser.detect_vba_macros():
                # マクロが検出された場合
                modules = []
                
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    module_info = {
                        "module_name": vba_filename,
                        "code": vba_code
                    }
                    modules.append(module_info)
                
                result = {
                    "found_macros": True,
                    "modules": modules,
                    "summary": f"Found {len(modules)} VBA modules"
                }
            else:
                # マクロが検出されなかった場合
                result = {
                    "found_macros": False,
                    "modules": [],
                    "summary": "No VBA macros found in the file"
                }
                
            vba_parser.close()
            
        except Exception as e:
            result = {
                "found_macros": False,
                "modules": [],
                "error": str(e),
                "summary": f"Error while processing the file: {str(e)}"
            }
            
        return result
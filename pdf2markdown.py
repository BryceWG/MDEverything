import requests

def pdf_to_markdown(pdf_file_path, app_id=None, secret_code=None, **kwargs):
    """
    将PDF文件转换为Markdown格式。

    :param pdf_file_path: PDF文件路径
    :param app_id: API应用ID
    :param secret_code: API密钥
    :param kwargs: 其他可选参数
    :return: 转换后的Markdown内容
    """
    url = "https://api.textin.com/ai/service/v1/pdf_to_markdown"
    
    headers = {
        "x-ti-app-id": app_id or "your_app_id_here",
        "x-ti-secret-code": secret_code or "your_secret_code_here"
    }
    
    params = {
        "apply_document_tree": kwargs.get('apply_document_tree', 1),
        "markdown_details": kwargs.get('markdown_details', 1),
        "table_flavor": kwargs.get('table_flavor', "md"),
        "get_image": kwargs.get('get_image', "none"),
        "dpi": kwargs.get('dpi', 144)
    }
    
    try:
        with open(pdf_file_path, "rb") as file:
            pdf_content = file.read()
        
        response = requests.post(url, headers=headers, params=params, data=pdf_content)
        response.raise_for_status()
        
        result = response.json()
        if result["code"] == 200:
            return result["result"]["markdown"]
        else:
            raise APIError(f"API错误: {result['message']}")
    except requests.RequestException as e:
        raise ConnectionError(f"HTTP请求错误: {str(e)}")
    except IOError as e:
        raise FileNotFoundError(f"无法读取PDF文件: {str(e)}")
    except Exception as e:
        raise ConversionError(f"PDF转换失败: {str(e)}")

class APIError(Exception):
    """API返回错误时抛出的异常"""
    pass

class ConversionError(Exception):
    """PDF转换过程中发生错误时抛出的异常"""
    pass

# 使用示例
if __name__ == "__main__":
    pdf_file_path = "path/to/your/pdf/file.pdf"
    app_id = "your_app_id"
    secret_code = "your_secret_code"
    
    try:
        markdown_content = pdf_to_markdown(pdf_file_path, app_id, secret_code)
        print("转换成功!")
        print("Markdown内容:")
        print(markdown_content)
    except (APIError, ConnectionError, FileNotFoundError, ConversionError) as e:
        print(f"转换失败: {str(e)}")
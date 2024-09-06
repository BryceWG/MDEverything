import os
import subprocess
import sys

def pptx_to_markdown(input_path, output_dir, **options):
    """
    将PowerPoint文件转换为Markdown格式。

    :param input_path: PowerPoint文件的路径
    :param output_dir: 输出目录
    :param options: 其他转换选项
    :return: 生成的Markdown文件路径
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_file = os.path.join(output_dir, f"{base_name}.md")
    img_folder = os.path.join(output_dir, f"{base_name}_img")
    
    # 构建pptx2md命令
    cmd = [sys.executable, "-m", "pptx2md", input_path,
           "--image-dir", img_folder,
           "-o", output_file]
    
    # 添加可选参数
    cmd.extend(build_optional_args(options))
    
    # 执行pptx2md命令
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"PPTX转换失败: {e.stderr}")
    
    # 检查输出文件是否存在
    if os.path.exists(output_file):
        return output_file
    else:
        raise FileNotFoundError(f"转换后的文件 {output_file} 不存在")

def build_optional_args(options):
    """构建pptx2md的可选参数列表"""
    args = []
    if options.get('image_width'):
        args.extend(["--image-width", str(options['image_width'])])
    if options.get('disable_image'):
        args.append("--disable-image")
    if options.get('disable_escaping'):
        args.append("--disable-escaping")
    if options.get('disable_notes'):
        args.append("--disable-notes")
    if options.get('disable_wmf', True):
        args.append("--disable-wmf")
    if options.get('disable_color'):
        args.append("--disable-color")
    if options.get('enable_slides'):
        args.append("--enable-slides")
    if options.get('min_block_size'):
        args.extend(["--min-block-size", str(options['min_block_size'])])
    
    # 处理输出格式
    output_format = options.get('output_format', 'markdown')
    if output_format in ['wiki', 'mdk', 'qmd']:
        args.append(f"--{output_format}")
    
    return args
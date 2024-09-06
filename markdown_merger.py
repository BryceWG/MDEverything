import os

def merge_markdown_files(selected_files, output_dir):
    merged_content = []
    for filename in selected_files:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()
        merged_content.append(f"# [{os.path.basename(filename)}]\n## Content\n{content}\n---\n")

    output_file = os.path.join(output_dir, "merged_markdown.md")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("\n".join(merged_content))

    return output_file
def replace_tabs_with_spaces(file_name):
    try:
        # 打开文件并读取内容
        with open(file_name, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 替换制表符为4个空格
        modified_content = content.replace('\t', '    ')
        
        # 将修改后的内容写回文件
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(modified_content)
        
        print(f"文件 {file_name} 中的制表符已成功替换为4个空格。")
    except FileNotFoundError:
        print(f"错误：文件 {file_name} 未找到，请确保文件名和后缀正确无误。")
    except Exception as e:
        print(f"发生错误：{e}")

# 主程序
if __name__ == "__main__":
    # 提示用户输入文件名
    file_name = input("请输入文件名称（包括后缀）：")
    replace_tabs_with_spaces(file_name)
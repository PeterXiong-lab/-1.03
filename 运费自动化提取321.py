import pandas as pd
import re

file_path = '321.xlsx'

def extract_to_notepad_format():
    print("--- 正在按指定记事本格式进行一键自动化提取 ---")
    try:
        df = pd.read_excel(file_path, header=None)
        
        output_blocks = []
        state = "SEARCH_HEADER"
        
        block_order = None
        block_city = None
        block_express_name = "韵达"  # 默认值
        
        order_pattern = re.compile(r'(PO-\d{8}-\d{8})')
        provinces = ['湖南', '湖北', '广东', '山东', '河南', '广西', '重庆', '四川', 
                     '江苏', '浙江', '安徽', '福建', '江西', '北京', '天津', '上海', 
                     '河北', '山西', '辽宁', '吉林', '黑龙江', '陕西', '甘肃', '青海', 
                     '贵州', '云南', '海南', '内蒙古', '西藏', '宁夏', '新疆']
        
        for index, row in df.iterrows():
            row_cells = [str(cell).strip() for cell in row if pd.notnull(cell) and str(cell).strip() != '']
            row_str = " ".join(row_cells)
            
            # 1. 捕捉块开头：获取订单号、提取纯城市名
            if "总" in row_str and "件" in row_str:
                order_match = order_pattern.search(row_str)
                if order_match:
                    block_order = order_match.group(1)
                    prefix_part = row_str.split("总")[0].strip()
                    for p in provinces:
                        if prefix_part.startswith(p):
                            prefix_part = prefix_part[len(p):].strip()
                    
                    # 仅保留前两个字作为地级市名称
                    block_city = prefix_part[:2] if len(prefix_part) >= 2 else prefix_part
                    state = "SEARCH_COLUMNS"
                    continue
            
            # 2. 捕捉表头行：动态识别当前栏目的快递名称（如“韵达”、“极兔”）
            if state == "SEARCH_COLUMNS":
                for cell in row_cells:
                    if "费用" in cell and "壹米" not in cell and "物流" not in cell:
                        block_express_name = cell.replace("费用", "")
                state = "SEARCH_REAL_PAY"
                continue
            
            # 3. 捕捉实付行：提取费用数据并按要求换行组装
            if state == "SEARCH_REAL_PAY" and "实付" in row_str:
                valid_numbers = []
                for cell in row:
                    if pd.notnull(cell):
                        try:
                            valid_numbers.append(float(str(cell).strip()))
                        except ValueError:
                            pass
                
                kuaidi_val = valid_numbers[0] if len(valid_numbers) >= 1 else 0
                wuliu_val = valid_numbers[1] if len(valid_numbers) >= 2 else 0
                
                # 格式化数字，去掉末尾无效的0
                kd_str = f"{kuaidi_val:.2f}".rstrip('0').rstrip('.') if kuaidi_val > 0 else "0"
                wl_str = f"{wuliu_val:.5f}".rstrip('0').rstrip('.') if wuliu_val > 0 else "0"
                
                # 严格按照要求的换行格式拼接
                block_text = (
                    f"{block_order}\n"
                    f"辛苦下单邮费链接备注：天猫美团{block_city}仓{block_express_name}\n"
                    f"快递{kd_str} 物流{wl_str}"
                )
                output_blocks.append(block_text)
                state = "SEARCH_HEADER"

        # 4. 写入记事本文件
        if output_blocks:
            final_text = "\n\n".join(output_blocks)
            with open("物流费用提取结果.txt", "w", encoding="utf-8") as f:
                f.write(final_text)
            print("\n🎉【大成功】记事本格式文本已生成在：'物流费用提取结果.txt'")
        else:
            print("\n❌ 未能成功提取，请检查Excel文件内容。")
            
    except Exception as e:
        print(f"\n运行出错: {e}")

if __name__ == '__main__':
    extract_to_notepad_format()

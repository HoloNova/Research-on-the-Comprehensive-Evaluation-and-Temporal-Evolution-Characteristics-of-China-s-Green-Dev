import pandas as pd
import os

folders = ['国民消费水平', '科技', '能源', '资源与环境']

for folder in folders:
    folder_path = os.path.join('.', folder)
    if os.path.exists(folder_path):
        print(f"\n{'='*80}")
        print(f"文件夹: {folder}")
        print('='*80)
        
        original_files = [f for f in os.listdir(folder_path) 
                         if f.endswith('.xlsx') and '已处理' not in f and not f.startswith('~$')]
        
        for file in original_files:
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path, header=None)
                print(f"\n文件: {file}")
                print(f"数据形状: {df.shape}")
                
                if len(df) > 3:
                    indicators = df.iloc[3:, 0].dropna().tolist()
                    indicators = [str(i) for i in indicators if '数据来源' not in str(i)]
                    print(f"包含指标: {indicators[:5]}")
                    if len(indicators) > 5:
                        print(f"  ... 还有 {len(indicators)-5} 个指标")
            except Exception as e:
                print(f"读取错误: {e}")

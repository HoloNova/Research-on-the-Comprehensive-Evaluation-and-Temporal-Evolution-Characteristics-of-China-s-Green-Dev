import sys
from data_processor import DataProcessor

def main():
    print("="*100)
    print("Excel数据系统化处理工具 - 运行模式选择")
    print("="*100)
    print("\n请选择运行模式:")
    print("1. 完整交互模式（推荐）- 每步暂停确认，显示可视化图表")
    print("2. 快速处理模式 - 自动处理，无暂停，无可视化")
    print("3. 退出")
    
    choice = input("\n请输入选择 (1/2/3): ").strip()
    
    if choice == '1':
        print("\n启动完整交互模式...")
        processor = DataProcessor()
        processor.process_all_files()
    elif choice == '2':
        print("\n启动快速处理模式...")
        from data_processor_fast import DataProcessorFast
        processor = DataProcessorFast()
        processor.process_all_files()
    elif choice == '3':
        print("\n退出程序")
        sys.exit(0)
    else:
        print("\n无效选择，退出程序")
        sys.exit(1)

if __name__ == "__main__":
    main()
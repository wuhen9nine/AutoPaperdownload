import csv
import os
import re
import logging
import traceback
from typing import List, Optional, Dict
from datetime import datetime

# 配置参数
OUTPUT_FOLDER = r"D:\Paperdownload\RSS"
CSV_FILE = r"D:\Paperdownload\LAsPaperDoi.csv"
LOG_FILE = r"D:\Paperdownload\doi_extractor.log"  # 日志文件路径

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def extract_strict_dois(text: str) -> List[str]:
    """
    严格提取 doi: 10.xxxx/xxxx 格式的DOI
    返回格式：["10.1088/2057-1976/adf8ee", ...]
    去除</dc:identifier>等XML/HTML标签后缀和doi:前缀
    确保DOI结尾不包含句点(.)
    """
    # 严格匹配 doi: 10.xxxx/xxxx 格式
    strict_doi_pattern = r'\bdoi:\s*(10\.[0-9]{4,}(?:\.[0-9]+)*/[^\s<;,)]+)'
    matches = re.findall(strict_doi_pattern, text, re.IGNORECASE)
    
    # 清理结果
    clean_dois = []
    for doi in matches:
        # 移除XML/HTML标签
        clean_doi = re.sub(r'<[^>]+>', '', doi)
        # 移除末尾的标点，特别注意去除句点
        clean_doi = re.sub(r'[^0-9a-zA-Z./-]+$', '', clean_doi)
        clean_doi = clean_doi.rstrip('.')
        clean_dois.append(clean_doi.lower())
    
    return clean_dois

def update_doi_csv(dois: List[str]) -> Optional[Dict]:
    """更新CSV文件中的DOI记录，只写入不存在的纯DOI号"""
    try:
        existing_dois = set()
        # 读取现有CSV文件中的所有DOI
        if os.path.exists(CSV_FILE):
            with open(CSV_FILE, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)  # 跳过标题行
                existing_dois = {row[0].lower().strip() for row in reader if row and row[0]}
        
        # 首先对当前提取的DOI列表去重
        unique_dois = list({doi.lower().strip() for doi in dois})
        if len(unique_dois) < len(dois):
            logger.info(f"注意：当前提取的DOI中有 {len(dois)-len(unique_dois)} 个重复值已被过滤")
        
        # 筛选出新DOI（不包含在existing_dois中的）
        new_dois = [doi for doi in unique_dois 
                   if doi not in existing_dois]
        
        if new_dois:
            # 写入新DOI（只包含DOI号，不包含"doi:"前缀）
            with open(CSV_FILE, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for doi in new_dois:
                    writer.writerow([doi, '', '', '', '','', '', ''])
            logger.info(f"已添加 {len(new_dois)} 个新DOI到CSV文件")
            logger.info("新增DOI列表:")
            for doi in new_dois:
                logger.info(f"  - {doi}")
        else:
            logger.info("没有发现新的DOI需要添加")
            
        # 返回统计信息
        return {
            "total_extracted": len(dois),
            "duplicates_in_current": len(dois) - len(unique_dois),
            "new_dois_added": len(new_dois),
            "existing_dois": len(existing_dois)
        }
    except Exception as e:
        logger.error(f"更新CSV文件失败: {str(e)}")
        return None

def initialize_csv():
    """初始化CSV文件（如果不存在）"""
    if not os.path.exists(CSV_FILE):
        os.makedirs(os.path.dirname(CSV_FILE), exist_ok=True)
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(['DOI', 'DownloadStatus', 'Filename', "URL", "DownloadURL","SIDownloadStatus","SIFilename","HTMLFilename"])
        logger.info(f"已创建新的CSV文件: {CSV_FILE}")

def get_latest_rss_file() -> Optional[str]:
    """获取最新的RSS文件"""
    try:
        # 确保目录存在
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        
        # 获取所有以"PubMed_RSS_"开头的txt文件
        rss_files = [f for f in os.listdir(OUTPUT_FOLDER) 
                    if f.startswith("PubMed_RSS_") and f.endswith(".txt")]
        
        if not rss_files:
            logger.warning(f"在目录 {OUTPUT_FOLDER} 中未找到RSS文件")
            return None
        
        # 按修改时间排序获取最新文件
        rss_files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
        latest_file = os.path.join(OUTPUT_FOLDER, rss_files[0])
        logger.info(f"找到最新的RSS文件: {latest_file}")
        return latest_file
    except Exception as e:
        logger.error(f"获取最新RSS文件失败: {str(e)}")
        return None

def process_rss_file(file_path: str) -> bool:
    """处理RSS文件并提取DOI"""
    try:
        # 读取文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if not content:
            logger.warning(f"文件内容为空: {file_path}")
            return False
        
        # 提取DOI
        dois = extract_strict_dois(content)
        if not dois:
            logger.warning(f"未在文件中找到任何DOI: {file_path}")
            return False
        
        logger.info(f"从文件中提取到 {len(dois)} 个DOI")
        
        # 更新CSV文件
        stats = update_doi_csv(dois)
        if stats:
            logger.info("\nDOI统计信息:")
            logger.info(f"- 本次提取DOI总数: {stats['total_extracted']}")
            logger.info(f"- 本次提取中的重复DOI: {stats['duplicates_in_current']}")
            logger.info(f"- 新增DOI数量: {stats['new_dois_added']}")
            logger.info(f"- 已有DOI总数: {stats['existing_dois']}")
            return True
        return False
    except Exception as e:
        logger.error(f"处理RSS文件失败: {str(e)}")
        return False

def main():
    logger.info("=== PubMed RSS DOI提取程序 ===")
    logger.info(f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # 初始化CSV文件
        initialize_csv()
        
        # 获取最新的RSS文件
        rss_file = get_latest_rss_file()
        if not rss_file:
            logger.error("未找到有效的RSS文件，程序终止")
            return
        
        # 处理RSS文件
        if process_rss_file(rss_file):
            logger.info("\nDOI提取程序执行完成")
        else:
            logger.warning("DOI提取过程中出现问题")
    except Exception as e:
        logger.error(f"程序执行过程中发生错误: {str(e)}")
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    # 创建必要的目录
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    main()
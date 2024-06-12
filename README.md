import base64
import os
import requests
from bs4 import BeautifulSoup
from docx import Document
import tkinter as tk
from tkinter import messagebox

def fetch_page_content(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.text
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP錯誤: {http_err}")
    except Exception as err:
        print(f"其他錯誤: {err}")

def parse_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    title = soup.title.string if soup.title else '無標題'
    texts = list(soup.stripped_strings)
    links = [a['href'] for a in soup.find_all('a', href=True)]
    images = [img['src'] for img in soup.find_all('img', src=True)]
    
    return title, texts, links, images

def save_to_word(title, texts, links, images, filename):
    doc = Document()
    doc.add_heading(title, 0)
    
    # 添加文本內容
    doc.add_heading('文本內容', level=1)
    for text in texts:
        doc.add_paragraph(text)
    
    # 添加連結
    doc.add_heading('連結', level=1)
    for link in links:
        doc.add_paragraph("連結：")
        doc.add_paragraph(link)
    
    # 添加圖片
    doc.add_heading('圖片', level=1)
    for image_data in images:
        try:
            # 解析 Base64 編碼的圖片數據
            image_type, image_data = image_data.split(';base64,')
            image_extension = image_type.split('/')[-1]
            image_bytes = base64.b64decode(image_data)
        
            # 將圖片保存到本地文件
            image_filename = f"image_{images.index(image_data)}.{image_extension}"
            with open(image_filename, 'wb') as f:
                f.write(image_bytes)
        
            # 將圖片插入到 Word 文件中
            doc.add_paragraph("圖片：")
            doc.add_picture(image_filename)
        
            # 刪除暫存的圖片文件
            os.remove(image_filename)
        except Exception as e:
            print(f"處理圖片時出錯: {e}")
    
    doc.save(filename)

def start_scraping():
    url = url_entry.get()
    filename = filename_entry.get()
    
    if not url:
        messagebox.showwarning("輸入錯誤", "請輸入URL。")
        return
    
    if not filename:
        messagebox.showwarning("輸入錯誤", "請輸入文件名。")
        return
    
    if not filename.endswith('.docx'):
        filename += '.docx'
    
    html_content = fetch_page_content(url)
    if not html_content:
        messagebox.showerror("獲取錯誤", "無法獲取網頁內容。")
        return
    
    title, texts, links, images = parse_html(html_content)
    save_to_word(title, texts, links, images, filename)
    messagebox.showinfo("成功", f"內容已保存到 {filename}")

# 創建圖形化介面
root = tk.Tk()
root.title("網頁爬蟲")

tk.Label(root, text="請輸入URL:").grid(row=0, column=0, padx=10, pady=10)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="請輸入文件名:").grid(row=1, column=0, padx=10, pady=10)
filename_entry = tk.Entry(root, width=50)
filename_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="開始爬取", command=start_scraping).grid(row=2, column=0, columnspan=2, pady=20)

root.mainloop()

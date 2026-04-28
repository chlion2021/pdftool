import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
import os
from typing import List, Tuple
import threading

class PDFFileListFrame(ttk.Frame):
    """PDF文件列表控件"""
    def __init__(self, parent):
        super().__init__(parent)
        
        # 创建Treeview作为列表
        self.tree = ttk.Treeview(self, columns=('文件名', '页数', '文件路径'), show='headings')
        
        # 设置列
        self.tree.heading('文件名', text='文件名')
        self.tree.heading('页数', text='页数')
        self.tree.heading('文件路径', text='文件路径')
        
        self.tree.column('文件名', width=300)
        self.tree.column('页数', width=80)
        self.tree.column('文件路径', width=400)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def add_file(self, file_path: str, page_count: int):
        """添加文件到列表"""
        filename = os.path.basename(file_path)
        self.tree.insert('', tk.END, values=(filename, page_count, file_path))
        
    def remove_selected(self):
        """删除选中的项目"""
        selected = self.tree.selection()
        for item in selected:
            self.tree.delete(item)
            
    def get_all_files(self) -> List[str]:
        """获取所有文件路径"""
        files = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            files.append(values[2])  # 文件路径在第三列
        return files
        
    def get_selected(self):
        """获取选中的项目"""
        return self.tree.selection()
        
    def move_up(self):
        """上移选中项"""
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        prev = self.tree.prev(item)
        if prev:
            # 获取当前和上一个项目的数据
            curr_values = self.tree.item(item)['values']
            prev_values = self.tree.item(prev)['values']
            # 交换
            self.tree.item(item, values=prev_values)
            self.tree.item(prev, values=curr_values)
            self.tree.selection_set(prev)
            
    def move_down(self):
        """下移选中项"""
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        next_item = self.tree.next(item)
        if next_item:
            curr_values = self.tree.item(item)['values']
            next_values = self.tree.item(next_item)['values']
            self.tree.item(item, values=next_values)
            self.tree.item(next_item, values=curr_values)
            self.tree.selection_set(next_item)
            
    def clear_all(self):
        """清空所有"""
        for item in self.tree.get_children():
            self.tree.delete(item)

class PDFSplitDialog(tk.Toplevel):
    """PDF分割对话框"""
    def __init__(self, parent, pdf_path: str, total_pages: int):
        super().__init__(parent)
        self.title("分割PDF - 输入页码后，分割前务必先按【添加范围】按钮")
        self.geometry("950x750")
        self.pdf_path = pdf_path
        self.total_pages = total_pages
        self.split_ranges = []
        self.result = None
        
        self.init_ui()
        self.transient(parent)
        self.grab_set()
        
    def init_ui(self):
        # 文件信息
        info_text = f"文件: {os.path.basename(self.pdf_path)}\n总页数: {self.total_pages}"
        info_label = ttk.Label(self, text=info_text, font=('Arial', 12, 'bold'))
        info_label.pack(pady=10)
        
        # 快速分割按钮
        quick_frame = ttk.Frame(self)
        quick_frame.pack(pady=5)
        ttk.Label(quick_frame, text="快速分割:").pack(side=tk.LEFT, padx=5)
        ttk.Button(quick_frame, text="逐页分割", command=self.on_every_page_split).pack(side=tk.LEFT, padx=5)
        
        # 分割线
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)
        
        # 分割范围输入
        range_frame = ttk.Frame(self)
        range_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(range_frame, text="自定义分割范围 (例如: 1-5,8,10-15):").pack(side=tk.LEFT, padx=5)
        self.range_text = ttk.Entry(range_frame, width=40)
        self.range_text.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(range_frame, text="添加范围", command=self.on_add_range).pack(side=tk.LEFT, padx=5)
        
        # 分割范围列表
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建Treeview
        self.range_tree = ttk.Treeview(list_frame, columns=('页码范围', '页数', '输出文件名'), show='headings')
        self.range_tree.heading('页码范围', text='页码范围')
        self.range_tree.heading('页数', text='页数')
        self.range_tree.heading('输出文件名', text='输出文件名')
        
        self.range_tree.column('页码范围', width=280)
        self.range_tree.column('页数', width=100)
        self.range_tree.column('输出文件名', width=300)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.range_tree.yview)
        self.range_tree.configure(yscrollcommand=scrollbar.set)
        
        self.range_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 操作按钮
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="删除选中", command=self.on_remove_range).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清空所有", command=self.on_clear_all).pack(side=tk.LEFT, padx=5)
        
        # 底部按钮
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(pady=10)
        
        ttk.Button(bottom_frame, text="开始分割", command=self.on_split).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=self.on_cancel).pack(side=tk.LEFT, padx=10)
        
        # 提示
        tip_text = "提示：\n• 单个页码: 5\n• 页码范围: 1-10\n• 多个范围: 1-5,8,10-15\n• 页码从1开始"
        tip_label = ttk.Label(self, text=tip_text, foreground='gray')
        tip_label.pack(pady=10)
        
    def parse_range(self, range_str: str) -> List[int]:
        """解析页码范围"""
        pages = set()
        parts = range_str.replace(' ', '').split(',')
        
        for part in parts:
            if not part:
                continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start < 1 or end > self.total_pages or start > end:
                        raise ValueError(f"无效范围: {part}")
                    pages.update(range(start, end + 1))
                except ValueError:
                    raise ValueError(f"无效范围格式: {part}")
            else:
                try:
                    page = int(part)
                    if page < 1 or page > self.total_pages:
                        raise ValueError(f"无效页码: {page}")
                    pages.add(page)
                except ValueError:
                    raise ValueError(f"无效页码格式: {part}")
        
        if not pages:
            raise ValueError("未指定有效的页码范围")
        
        return sorted(pages)
    
    def on_every_page_split(self):
        """逐页分割"""
        # 清空现有范围
        for item in self.range_tree.get_children():
            self.range_tree.delete(item)
        self.split_ranges.clear()
        
        # 为每一页创建一个范围
        for i in range(1, self.total_pages + 1):
            range_str = str(i)
            pages = [i]
            output_name = f"page_{i:03d}.pdf"
            
            self.range_tree.insert('', tk.END, values=(range_str, "1", output_name))
            self.split_ranges.append((pages, range_str, output_name))
        
        messagebox.showinfo("提示", f"已添加 {self.total_pages} 个分割范围，每页一个文件")
        
    def on_add_range(self):
        """添加分割范围"""
        range_str = self.range_text.get().strip()
        if not range_str:
            messagebox.showwarning("提示", "请输入分割范围")
            return
            
        try:
            pages = self.parse_range(range_str)
            output_name = f"split_{len(self.split_ranges) + 1}.pdf"
            
            self.range_tree.insert('', tk.END, values=(range_str, len(pages), output_name))
            self.split_ranges.append((pages, range_str, output_name))
            self.range_text.delete(0, tk.END)
            
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            
    def on_remove_range(self):
        """删除选中的分割范围"""
        selected = self.range_tree.selection()
        if selected:
            # 获取选中项在列表中的索引
            items = self.range_tree.get_children()
            index = items.index(selected[0])
            self.range_tree.delete(selected[0])
            self.split_ranges.pop(index)
            # 重新编号
            for i, item in enumerate(items[1:] if index == 0 else items):
                if item in self.range_tree.get_children():
                    values = list(self.range_tree.item(item)['values'])
                    if len(values) == 3:
                        values[2] = f"split_{i + 1}.pdf"
                        self.range_tree.item(item, values=values)
                        if i < len(self.split_ranges):
                            self.split_ranges[i] = (self.split_ranges[i][0], self.split_ranges[i][1], values[2])
                
    def on_clear_all(self):
        """清空所有分割范围"""
        if self.split_ranges:
            if messagebox.askyesno("确认", "确定要清空所有分割范围吗？"):
                for item in self.range_tree.get_children():
                    self.range_tree.delete(item)
                self.split_ranges.clear()
            
    def on_split(self):
        """执行分割"""
        if not self.split_ranges:
            messagebox.showwarning("提示", "请至少添加一个分割范围")
            return
        self.result = self.split_ranges
        self.destroy()
        
    def on_cancel(self):
        self.result = None
        self.destroy()
        
    def get_split_info(self):
        return self.result

class PDFMergerSplitter:
    """PDF合并分割主窗口"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF合并分割工具 作者:陈宏宇 微信：chenhongyusnow 20260119 赠人玫瑰，手有余香")
        self.root.geometry("1000x700")
        
        self.init_ui()
        self.center_window()
        
    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def init_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建标签页
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 合并标签页
        merge_frame = self.create_merge_panel()
        notebook.add(merge_frame, text="PDF合并")
        
        # 分割标签页
        split_frame = self.create_split_panel()
        notebook.add(split_frame, text="PDF分割")
        
        # 创建状态栏
        self.statusbar = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        
    def create_merge_panel(self):
        """创建合并面板"""
        panel = ttk.Frame()
        
        # 文件列表
        self.merge_list = PDFFileListFrame(panel)
        self.merge_list.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 按钮面板
        btn_frame = ttk.Frame(panel)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="添加PDF文件", command=self.on_add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除选中", command=self.on_remove_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空列表", command=self.on_clear_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="上移", command=self.on_move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下移", command=self.on_move_down).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame, text="合并PDF", command=self.on_merge_pdfs).pack(side=tk.RIGHT, padx=2)
        
        return panel
        
    def create_split_panel(self):
        """创建分割面板"""
        panel = ttk.Frame()
        
        # 文件选择
        file_frame = ttk.Frame(panel)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="PDF文件:").pack(side=tk.LEFT, padx=5)
        self.split_file_path = ttk.Entry(file_frame, state='readonly')
        self.split_file_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(file_frame, text="选择PDF文件", command=self.on_select_split_file).pack(side=tk.LEFT, padx=5)
        
        # 文件信息显示
        self.split_info = scrolledtext.ScrolledText(panel, height=5, state='disabled')
        self.split_info.pack(fill=tk.X, pady=5, padx=5)
        
        # 分割按钮面板
        btn_frame = ttk.Frame(panel)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="自定义分割", command=self.on_custom_split, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="逐页分割", command=self.on_every_page_split, width=15).pack(side=tk.LEFT, padx=10)
        
        return panel
        
    def get_pdf_page_count(self, file_path: str) -> int:
        """获取PDF页数"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return len(pdf_reader.pages)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
            return 0
            
    def on_add_files(self):
        """添加PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )
        added_count = 0
        for path in files:
            if path.lower().endswith('.pdf'):
                existing_files = self.merge_list.get_all_files()
                if path not in existing_files:
                    page_count = self.get_pdf_page_count(path)
                    if page_count > 0:
                        self.merge_list.add_file(path, page_count)
                        added_count += 1
        if added_count > 0:
            self.statusbar.config(text=f"已添加 {added_count} 个文件")
                
    def on_remove_files(self):
        """删除选中的文件"""
        count = len(self.merge_list.get_all_files())
        self.merge_list.remove_selected()
        new_count = len(self.merge_list.get_all_files())
        self.statusbar.config(text=f"已删除 {count - new_count} 个文件")
        
    def on_clear_files(self):
        """清空所有文件"""
        if len(self.merge_list.get_all_files()) > 0:
            if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
                self.merge_list.clear_all()
                self.statusbar.config(text="已清空列表")
        
    def on_move_up(self):
        """上移文件"""
        self.merge_list.move_up()
        self.statusbar.config(text="已上移文件")
            
    def on_move_down(self):
        """下移文件"""
        self.merge_list.move_down()
        self.statusbar.config(text="已下移文件")
            
    def on_merge_pdfs(self):
        """合并PDF文件"""
        files = self.merge_list.get_all_files()
        if len(files) < 2:
            messagebox.showwarning("提示", "请至少添加2个PDF文件进行合并")
            return
            
        output_path = filedialog.asksaveasfilename(
            title="保存合并后的PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if output_path:
            self.statusbar.config(text="正在合并PDF文件...")
            self.merge_pdfs_thread(files, output_path)
                
    def merge_pdfs_thread(self, files: List[str], output_path: str):
        """在线程中合并PDF"""
        def merge():
            try:
                pdf_merger = PyPDF2.PdfMerger()
                for file_path in files:
                    pdf_merger.append(file_path)
                pdf_merger.write(output_path)
                pdf_merger.close()
                self.root.after(0, self.show_merge_result, True, output_path)
            except Exception as e:
                self.root.after(0, self.show_merge_result, False, str(e))
                
        thread = threading.Thread(target=merge)
        thread.daemon = True
        thread.start()
        
    def show_merge_result(self, success: bool, message: str):
        """显示合并结果"""
        if success:
            self.statusbar.config(text="合并完成")
            messagebox.showinfo("成功", f"PDF合并成功！\n保存位置: {message}")
        else:
            self.statusbar.config(text="合并失败")
            messagebox.showerror("错误", f"合并失败: {message}")
            
    def on_select_split_file(self):
        """选择要分割的PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择要分割的PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.split_file_path.config(state='normal')
            self.split_file_path.delete(0, tk.END)
            self.split_file_path.insert(0, file_path)
            self.split_file_path.config(state='readonly')
            
            page_count = self.get_pdf_page_count(file_path)
            if page_count > 0:
                file_size = os.path.getsize(file_path)
                if file_size < 1024 * 1024:
                    size_str = f"{file_size / 1024:.2f} KB"
                else:
                    size_str = f"{file_size / (1024 * 1024):.2f} MB"
                    
                self.split_info.config(state='normal')
                self.split_info.delete(1.0, tk.END)
                self.split_info.insert(tk.END, 
                    f"文件名: {os.path.basename(file_path)}\n"
                    f"文件路径: {file_path}\n"
                    f"总页数: {page_count} 页\n"
                    f"文件大小: {size_str}")
                self.split_info.config(state='disabled')
                self.statusbar.config(text=f"已选择文件: {os.path.basename(file_path)}")
                    
    def on_custom_split(self):
        """自定义分割"""
        file_path = self.split_file_path.get().strip()
        if not file_path:
            messagebox.showwarning("提示", "请先选择要分割的PDF文件")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在，请重新选择")
            return
            
        page_count = self.get_pdf_page_count(file_path)
        if page_count == 0:
            return
            
        dialog = PDFSplitDialog(self.root, file_path, page_count)
        self.root.wait_window(dialog)
        split_ranges = dialog.get_split_info()
        
        if split_ranges:
            output_dir = filedialog.askdirectory(title="选择保存目录")
            if output_dir:
                self.statusbar.config(text="正在分割PDF文件...")
                self.split_pdf_thread(file_path, split_ranges, output_dir)
        
    def on_every_page_split(self):
        """逐页分割"""
        file_path = self.split_file_path.get().strip()
        if not file_path:
            messagebox.showwarning("提示", "请先选择要分割的PDF文件")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在，请重新选择")
            return
            
        page_count = self.get_pdf_page_count(file_path)
        if page_count == 0:
            return
            
        if messagebox.askyesno("确认", f"确定要将PDF的 {page_count} 页逐页分割吗？\n将生成 {page_count} 个单独的PDF文件。"):
            output_dir = filedialog.askdirectory(title="选择保存目录")
            if output_dir:
                self.statusbar.config(text="正在逐页分割PDF文件...")
                split_ranges = []
                for i in range(1, page_count + 1):
                    pages = [i]
                    range_str = str(i)
                    output_name = f"page_{i:03d}.pdf"
                    split_ranges.append((pages, range_str, output_name))
                self.split_pdf_thread(file_path, split_ranges, output_dir)
        
    def split_pdf_thread(self, file_path: str, split_ranges: List, output_dir: str):
        """在线程中分割PDF"""
        def split():
            try:
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for i, (pages, range_str, output_name) in enumerate(split_ranges):
                        pdf_writer = PyPDF2.PdfWriter()
                        for page_num in pages:
                            pdf_writer.add_page(pdf_reader.pages[page_num - 1])
                        output_path = os.path.join(output_dir, output_name)
                        with open(output_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                self.root.after(0, self.show_split_result, True, output_dir, len(split_ranges))
            except Exception as e:
                self.root.after(0, self.show_split_result, False, str(e), 0)
                
        thread = threading.Thread(target=split)
        thread.daemon = True
        thread.start()
        
    def show_split_result(self, success: bool, message: str, count: int):
        """显示分割结果"""
        if success:
            self.statusbar.config(text="分割完成")
            messagebox.showinfo("成功", f"PDF分割成功！\n共生成 {count} 个文件\n保存位置: {message}")
        else:
            self.statusbar.config(text="分割失败")
            messagebox.showerror("错误", f"分割失败: {message}")
            
    def on_exit(self):
        """退出程序"""
        self.root.destroy()

def main():
    app = PDFMergerSplitter()
    app.root.mainloop()

if __name__ == "__main__":
    main()

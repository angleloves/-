import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
import subprocess
import time
import threading
from pathlib import Path

class DesktopLauncher:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("自动打开桌面软件")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f0f0')
        
        # 数据存储
        self.current_files = []  # 当前文件列表
        self.history_records = []  # 历史记录
        self.data_file = "launcher_data.json"
        
        # 初始化界面
        self.setup_ui()
        self.load_data()
        # 如果有历史记录，默认加载第一条
        self.load_first_history_on_startup()
        
    def setup_ui(self):
        # 主标题
        title_frame = tk.Frame(self.root, bg='#f0f0f0')
        title_frame.pack(fill='x', padx=20, pady=10)
        
        title_label = tk.Label(title_frame, text="自动打开桌面软件", 
                              font=('Microsoft YaHei', 24, 'bold'),
                              bg='#f0f0f0', fg='#2c3e50')
        title_label.pack()
        
        # 主容器
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 上半部分 - 左右分布
        top_frame = tk.Frame(main_frame, bg='#f0f0f0')
        top_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # 左侧面板 - 当前文件列表
        left_frame = tk.LabelFrame(top_frame, text="当前文件列表", 
                                  font=('Microsoft YaHei', 12, 'bold'),
                                  bg='#ffffff', fg='#2c3e50', relief='raised')
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 文件列表
        self.file_tree = ttk.Treeview(left_frame, columns=('序号', '文件路径', '延迟时间'), show='headings')
        self.file_tree.heading('序号', text='序号')
        self.file_tree.heading('文件路径', text='文件路径')
        self.file_tree.heading('延迟时间', text='延迟时间(秒)')
        
        self.file_tree.column('序号', width=60, anchor='center')
        self.file_tree.column('文件路径', width=400, anchor='w')
        self.file_tree.column('延迟时间', width=100, anchor='center')
        
        file_scroll = ttk.Scrollbar(left_frame, orient='vertical', command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=file_scroll.set)
        
        self.file_tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        file_scroll.pack(side='right', fill='y', pady=10)
        
        # 双击编辑事件
        self.file_tree.bind('<Double-1>', self.edit_file_item)
        
        # 右侧面板 - 历史记录
        right_frame = tk.LabelFrame(top_frame, text="历史记录", 
                                   font=('Microsoft YaHei', 12, 'bold'),
                                   bg='#ffffff', fg='#2c3e50', relief='raised')
        right_frame.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        # 历史记录列表
        self.history_tree = ttk.Treeview(right_frame, columns=('记录名称', '文件数量', '创建时间'), show='headings')
        self.history_tree.heading('记录名称', text='记录名称')
        self.history_tree.heading('文件数量', text='文件数量')
        self.history_tree.heading('创建时间', text='创建时间')
        
        self.history_tree.column('记录名称', width=200, anchor='w')
        self.history_tree.column('文件数量', width=80, anchor='center')
        self.history_tree.column('创建时间', width=150, anchor='center')
        
        history_scroll = ttk.Scrollbar(right_frame, orient='vertical', command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=history_scroll.set)
        
        self.history_tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        history_scroll.pack(side='right', fill='y', pady=10)
        
        # 双击编辑历史记录名称
        self.history_tree.bind('<Double-1>', self.edit_history_name)
        # 单击加载历史记录
        self.history_tree.bind('<<TreeviewSelect>>', self.load_history_on_click)
        
        # 下半部分 - 按钮区域
        bottom_frame = tk.Frame(main_frame, bg='#f0f0f0')
        bottom_frame.pack(fill='x', pady=(10, 0))
        
        # 第一行按钮 - 居中显示（合并所有按钮到一行）
        first_row_frame = tk.Frame(bottom_frame, bg='#f0f0f0')
        first_row_frame.pack(pady=5)
        
        btn_style = {'font': ('Microsoft YaHei', 10), 'width': 12, 'height': 1}
        
        tk.Button(first_row_frame, text="添加文件", command=self.add_file,
                 bg='#3498db', fg='white', **btn_style).pack(side='left', padx=5)
        tk.Button(first_row_frame, text="删除文件", command=self.delete_file,
                 bg='#e74c3c', fg='white', **btn_style).pack(side='left', padx=5)
        tk.Button(first_row_frame, text="上移", command=self.move_up,
                 bg='#9b59b6', fg='white', **btn_style).pack(side='left', padx=5)
        tk.Button(first_row_frame, text="下移", command=self.move_down,
                 bg='#9b59b6', fg='white', **btn_style).pack(side='left', padx=5)
        tk.Button(first_row_frame, text="保存记录", command=self.save_config,
                 bg='#27ae60', fg='white', **btn_style).pack(side='left', padx=5)
        tk.Button(first_row_frame, text="删除记录", command=self.delete_record,
                 bg='#e74c3c', fg='white', **btn_style).pack(side='left', padx=5)
        
        # 运行控制区域
        run_frame = tk.Frame(bottom_frame, bg='#f0f0f0')
        run_frame.pack(pady=10)
        
        # 运行后关闭软件选项
        self.close_after_run = tk.BooleanVar()
        tk.Checkbutton(run_frame, text="运行后关闭软件", variable=self.close_after_run,
                      font=('Microsoft YaHei', 10), bg='#f0f0f0').pack(side='left', padx=10)
        
        tk.Button(run_frame, text="开始运行", command=self.run_files,
                 bg='#f39c12', fg='white', font=('Microsoft YaHei', 12, 'bold'),
                 width=15, height=2).pack(side='left', padx=10)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                             relief='sunken', anchor='w', bg='#ecf0f1', fg='#2c3e50')
        status_bar.pack(side='bottom', fill='x')
    
    def load_first_history_on_startup(self):
        """启动时加载第一条历史记录"""
        if self.history_records:
            first_record = self.history_records[0]
            self.current_files = first_record['files'].copy()
            self.close_after_run.set(first_record.get('close_after_run', False))
            self.refresh_file_list()
            # 选中第一条历史记录
            if self.history_tree.get_children():
                first_item = self.history_tree.get_children()[0]
                self.history_tree.selection_set(first_item)
            self.status_var.set(f"已自动加载第一条历史记录: {first_record['name']}")
    
    def add_file(self):
        """添加文件"""
        file_path = filedialog.askopenfilename(
            title="选择要运行的文件",
            filetypes=[("所有文件", "*.*"), ("可执行文件", "*.exe"), ("批处理文件", "*.bat")]
        )
        
        if file_path:
            delay = simpledialog.askfloat("设置延迟时间", "请输入延迟执行时间(秒):", 
                                        minvalue=0, maxvalue=3600, initialvalue=0)
            if delay is not None:
                self.current_files.append({
                    'path': file_path,
                    'delay': delay,
                    'order': len(self.current_files) + 1
                })
                self.refresh_file_list()
                self.status_var.set(f"已添加文件: {os.path.basename(file_path)}")
    
    def delete_file(self):
        """删除选中的文件"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的文件")
            return
        
        item = selection[0]
        # 获取在TreeView中的索引位置
        children = self.file_tree.get_children()
        index = children.index(item)
        
        deleted_file = self.current_files.pop(index)
        # 重新排序
        for i, file_item in enumerate(self.current_files):
            file_item['order'] = i + 1
        self.refresh_file_list()
        self.status_var.set(f"已删除文件: {os.path.basename(deleted_file['path'])}")
    
    def edit_file_item(self, event):
        """编辑文件项"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        children = self.file_tree.get_children()
        index = children.index(item)
        current_file = self.current_files[index]
        
        # 编辑延迟时间
        new_delay = simpledialog.askfloat("修改延迟时间", 
                                        f"文件: {os.path.basename(current_file['path'])}\n当前延迟时间: {current_file['delay']}秒\n请输入新的延迟时间:",
                                        minvalue=0, maxvalue=3600, initialvalue=current_file['delay'])
        
        if new_delay is not None:
            current_file['delay'] = new_delay
            self.refresh_file_list()
            self.status_var.set("已更新延迟时间")
    
    def save_config(self):
        """保存当前配置到历史记录"""
        if not self.current_files:
            messagebox.showwarning("警告", "当前文件列表为空，无法保存")
            return
        
        record_name = simpledialog.askstring("保存配置", "请输入记录名称:")
        if record_name is None:  # 用户取消
            return
        
        if not record_name.strip():
            record_name = "历史记录"
        
        # 检查是否已存在同名记录
        for record in self.history_records:
            if record['name'] == record_name:
                if not messagebox.askyesno("记录已存在", f"记录'{record_name}'已存在，是否覆盖?"):
                    return
                self.history_records.remove(record)
                break
        
        # 创建新记录
        new_record = {
            'name': record_name,
            'files': self.current_files.copy(),
            'close_after_run': self.close_after_run.get(),
            'create_time': time.strftime("%Y-%m-%d %H:%M:%S")
        }
        
        self.history_records.append(new_record)
        self.refresh_history_list()
        self.save_data()
        self.status_var.set(f"已保存配置: {record_name}")
    
    def auto_save_default_record(self):
        """自动保存默认历史记录（仅在历史记录为空且有文件时）"""
        if not self.history_records and self.current_files:
            # 创建默认记录
            default_record = {
                'name': '默认历史记录',
                'files': self.current_files.copy(),
                'close_after_run': self.close_after_run.get(),
                'create_time': time.strftime("%Y-%m-%d %H:%M:%S")
            }
            
            self.history_records.append(default_record)
            self.refresh_history_list()
            self.save_data()
            self.status_var.set("已自动保存为默认历史记录")
    
    def load_history_on_click(self, event):
        """单击历史记录时加载配置"""
        selection = self.history_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        children = self.history_tree.get_children()
        index = children.index(item)
        
        if index < len(self.history_records):
            target_record = self.history_records[index]
            self.current_files = target_record['files'].copy()
            self.close_after_run.set(target_record.get('close_after_run', False))
            self.refresh_file_list()
            self.status_var.set(f"已加载配置: {target_record['name']}")
    
    def delete_record(self):
        """删除选中的历史记录"""
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的历史记录")
            return
        
        item = selection[0]
        # 获取在TreeView中的索引位置
        children = self.history_tree.get_children()
        index = children.index(item)
        
        record_name = self.history_records[index]['name']
        
        if messagebox.askyesno("确认删除", f"确定要删除历史记录: {record_name}?"):
            deleted_record = self.history_records.pop(index)
            self.refresh_history_list()
            self.save_data()
            self.status_var.set(f"已删除记录: {deleted_record['name']}")
    
    def move_up(self):
        """上移选中项"""
        # 优先处理文件列表
        file_selection = self.file_tree.selection()
        if file_selection:
            item = file_selection[0]
            children = self.file_tree.get_children()
            index = children.index(item)
            
            if index > 0:  # 不是第一个
                # 交换位置
                self.current_files[index], self.current_files[index - 1] = \
                    self.current_files[index - 1], self.current_files[index]
                
                # 重新排序
                for i, file_item in enumerate(self.current_files):
                    file_item['order'] = i + 1
                
                self.refresh_file_list()
                # 保持选中状态
                new_item = self.file_tree.get_children()[index - 1]
                self.file_tree.selection_set(new_item)
                self.status_var.set("已上移文件")
            else:
                self.status_var.set("已是第一个文件，无法上移")
            return
        
        # 处理历史记录列表
        history_selection = self.history_tree.selection()
        if history_selection:
            item = history_selection[0]
            children = self.history_tree.get_children()
            index = children.index(item)
            
            if index > 0:  # 不是第一个
                # 交换位置
                self.history_records[index], self.history_records[index - 1] = \
                    self.history_records[index - 1], self.history_records[index]
                
                self.refresh_history_list()
                self.save_data()
                # 保持选中状态
                new_item = self.history_tree.get_children()[index - 1]
                self.history_tree.selection_set(new_item)
                self.status_var.set("已上移历史记录")
            else:
                self.status_var.set("已是第一个记录，无法上移")
            return
        
        messagebox.showwarning("警告", "请先选择要移动的项目")
    
    def move_down(self):
        """下移选中项"""
        # 优先处理文件列表
        file_selection = self.file_tree.selection()
        if file_selection:
            item = file_selection[0]
            children = self.file_tree.get_children()
            index = children.index(item)
            
            if index < len(self.current_files) - 1:  # 不是最后一个
                # 交换位置
                self.current_files[index], self.current_files[index + 1] = \
                    self.current_files[index + 1], self.current_files[index]
                
                # 重新排序
                for i, file_item in enumerate(self.current_files):
                    file_item['order'] = i + 1
                
                self.refresh_file_list()
                # 保持选中状态
                new_item = self.file_tree.get_children()[index + 1]
                self.file_tree.selection_set(new_item)
                self.status_var.set("已下移文件")
            else:
                self.status_var.set("已是最后一个文件，无法下移")
            return
        
        # 处理历史记录列表
        history_selection = self.history_tree.selection()
        if history_selection:
            item = history_selection[0]
            children = self.history_tree.get_children()
            index = children.index(item)
            
            if index < len(self.history_records) - 1:  # 不是最后一个
                # 交换位置
                self.history_records[index], self.history_records[index + 1] = \
                    self.history_records[index + 1], self.history_records[index]
                
                self.refresh_history_list()
                self.save_data()
                # 保持选中状态
                new_item = self.history_tree.get_children()[index + 1]
                self.history_tree.selection_set(new_item)
                self.status_var.set("已下移历史记录")
            else:
                self.status_var.set("已是最后一个记录，无法下移")
            return
        
        messagebox.showwarning("警告", "请先选择要移动的项目")
    
    def edit_history_name(self, event):
        """编辑历史记录名称"""
        selection = self.history_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        children = self.history_tree.get_children()
        index = children.index(item)
        
        if index < len(self.history_records):
            old_name = self.history_records[index]['name']
            
            new_name = simpledialog.askstring("修改记录名称", f"当前名称: {old_name}\n请输入新名称:", initialvalue=old_name)
            
            if new_name and new_name != old_name:
                # 检查是否已存在同名记录
                for record in self.history_records:
                    if record['name'] == new_name:
                        messagebox.showerror("错误", "该名称已存在")
                        return
                
                # 更新记录名称
                self.history_records[index]['name'] = new_name
                self.refresh_history_list()
                self.save_data()
                self.status_var.set(f"已修改记录名称: {old_name} -> {new_name}")
    
    def run_files(self):
        """运行文件列表"""
        if not self.current_files:
            messagebox.showwarning("警告", "当前文件列表为空")
            return
        
        # 如果历史记录为空，自动保存默认记录
        self.auto_save_default_record()
        
        def run_in_thread():
            try:
                self.status_var.set("正在运行文件...")
                total_files = len(self.current_files)
                
                for i, file_item in enumerate(self.current_files, 1):
                    file_path = file_item['path']
                    delay = file_item['delay']
                    
                    # 先等待延迟时间
                    if delay > 0:
                        self.status_var.set(f"等待 {delay} 秒后运行第 {i} 个文件: {os.path.basename(file_path)}")
                        time.sleep(delay)
                    
                    self.status_var.set(f"正在运行第 {i}/{total_files} 个文件: {os.path.basename(file_path)}")
                    
                    # 检查文件是否存在
                    if not os.path.exists(file_path):
                        messagebox.showerror("错误", f"文件不存在: {file_path}")
                        continue
                    
                    try:
                        # 运行文件
                        if file_path.endswith('.exe') or file_path.endswith('.bat'):
                            subprocess.Popen([file_path])
                        else:
                            # 使用系统默认程序打开
                            os.startfile(file_path)
                        
                    except Exception as e:
                        messagebox.showerror("错误", f"无法运行文件 {file_path}:\n{str(e)}")
                
                self.status_var.set("所有文件运行完成")
                
                # 如果设置了运行后关闭软件
                if self.close_after_run.get():
                    time.sleep(1)  # 等待1秒后关闭
                    self.root.quit()
                    
            except Exception as e:
                messagebox.showerror("错误", f"运行过程中发生错误:\n{str(e)}")
                self.status_var.set("运行失败")
        
        # 在新线程中运行，避免阻塞UI
        thread = threading.Thread(target=run_in_thread)
        thread.daemon = True
        thread.start()
    
    def refresh_file_list(self):
        """刷新文件列表显示"""
        # 清空现有项目
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 添加文件项目
        for file_item in self.current_files:
            self.file_tree.insert('', 'end', values=(
                file_item['order'],
                file_item['path'],
                file_item['delay']
            ))
    
    def refresh_history_list(self):
        """刷新历史记录列表显示"""
        # 清空现有项目
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # 添加历史记录项目
        for record in self.history_records:
            self.history_tree.insert('', 'end', values=(
                record['name'],
                len(record['files']),
                record['create_time']
            ))
    
    def save_data(self):
        """保存数据到本地文件"""
        try:
            data = {
                'history_records': self.history_records
            }
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存数据失败:\n{str(e)}")
    
    def load_data(self):
        """从本地文件加载数据"""
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.history_records = data.get('history_records', [])
                    self.refresh_history_list()
                    self.status_var.set(f"已加载 {len(self.history_records)} 条历史记录")
            else:
                self.status_var.set("未找到历史数据文件")
        except Exception as e:
            messagebox.showerror("错误", f"加载数据失败:\n{str(e)}")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = DesktopLauncher()
    app.run()
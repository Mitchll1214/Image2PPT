import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import pythoncom
import math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image

class ImageToPPTConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("图片转PPT工具 - FJY鎏金专业版")
        
        # 设置窗口大小和位置
        self.root.geometry("600x500+400+200")
        self.root.resizable(False, False)
        
        # 鎏金风格配色
        self.bg_color = "#121212"  # 深黑背景
        self.fg_color = "#E6C229"   # 鎏金文字
        self.hl_color = "#FFD700"   # 高亮金色
        self.btn_bg = "#2A2A2A"    # 按钮背景
        self.entry_bg = "#252525"  # 输入框背景（调亮）
        self.entry_fg = "#FFFFFF"  # 输入框文字（白色更醒目）
        self.entry_bd = "#E6C229"  # 输入框边框（鎏金色）
        
        # 应用整体样式
        self.set_style()
        
        # 变量初始化
        self.source_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.images_per_page = tk.IntVar(value=6)
        self.show_filenames = tk.BooleanVar(value=True)
        self.add_prefix = tk.BooleanVar(value=True)
        self.prefix_text = tk.StringVar(value="FJY")
        
        # 创建UI
        self.create_widgets()
    
    def set_style(self):
        """设置鎏金风格的整体样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 进度条样式
        style.configure("TProgressbar",
                      thickness=20,
                      troughcolor=self.bg_color,
                      background=self.hl_color,
                      lightcolor=self.fg_color,
                      darkcolor=self.fg_color)
        
        # 按钮样式
        style.configure("TButton",
                      font=("微软雅黑", 10),
                      foreground=self.fg_color,
                      background=self.btn_bg,
                      bordercolor=self.hl_color,
                      focusthickness=3,
                      focuscolor=self.hl_color)
        style.map("TButton",
                foreground=[('active', self.hl_color)],
                background=[('active', self.btn_bg)])
        
        # 配置主窗口
        self.root.configure(bg=self.bg_color)
    
    def create_widgets(self):
        """创建鎏金风格的UI组件"""
        # 主框架
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        title = tk.Label(main_frame,
                       text="图片转PPT工具",
                       font=("微软雅黑", 18, "bold"),
                       fg=self.hl_color,
                       bg=self.bg_color)
        title.pack(pady=(0, 20))
        
        # 输入框框架
        input_frame = tk.Frame(main_frame, bg=self.bg_color)
        input_frame.pack(fill=tk.X, pady=5)
        
        # 源文件夹选择
        self.create_input_row(input_frame, 0, "图片源文件夹:", self.source_folder, self.browse_source)
        
        # 输出文件夹选择
        self.create_input_row(input_frame, 1, "PPT输出文件夹:", self.output_folder, self.browse_output)
        
        # 设置框架
        settings_frame = tk.Frame(main_frame, bg=self.bg_color)
        settings_frame.pack(fill=tk.X, pady=10)
        
        # 文件名前缀设置
        prefix_frame = tk.Frame(settings_frame, bg=self.bg_color)
        prefix_frame.pack(fill=tk.X, pady=5)
        
        tk.Checkbutton(prefix_frame,
                     text="添加文件名前缀:",
                     variable=self.add_prefix,
                     font=("微软雅黑", 10),
                     fg=self.fg_color,
                     bg=self.bg_color,
                     selectcolor=self.btn_bg,
                     activebackground=self.bg_color,
                     activeforeground=self.fg_color).pack(side=tk.LEFT)
        
        tk.Entry(prefix_frame,
               textvariable=self.prefix_text,
               width=10,
               font=("微软雅黑", 10),
               fg=self.entry_fg,
               bg=self.entry_bg,
               insertbackground=self.fg_color,
               relief=tk.SOLID,
               bd=1,
               highlightbackground=self.entry_bd,
               highlightthickness=1).pack(side=tk.LEFT, padx=5)
        
        # 每页图片数量
        count_frame = tk.Frame(settings_frame, bg=self.bg_color)
        count_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(count_frame,
               text="每页图片数量:",
               font=("微软雅黑", 10),
               fg=self.fg_color,
               bg=self.bg_color).pack(side=tk.LEFT)
        
        tk.Entry(count_frame,
               textvariable=self.images_per_page,
               width=5,
               font=("微软雅黑", 10),
               fg=self.entry_fg,
               bg=self.entry_bg,
               insertbackground=self.fg_color,
               relief=tk.SOLID,
               bd=1,
               highlightbackground=self.entry_bd,
               highlightthickness=1).pack(side=tk.LEFT, padx=5)
        
        # 显示文件名选项
        tk.Checkbutton(settings_frame,
                     text="在图片下方显示文件名",
                     variable=self.show_filenames,
                     font=("微软雅黑", 10),
                     fg=self.fg_color,
                     bg=self.bg_color,
                     selectcolor=self.btn_bg,
                     activebackground=self.bg_color,
                     activeforeground=self.fg_color).pack(anchor=tk.W, pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame,
                                     orient='horizontal',
                                     length=500,
                                     mode='determinate',
                                     style="TProgressbar")
        self.progress.pack(pady=15)
        
        # 转换按钮
        convert_btn = ttk.Button(main_frame,
                               text="开始转换",
                               command=self.convert_images_to_ppt,
                               style="TButton")
        convert_btn.pack(pady=10)
        
        # 版权信息
        copyright = tk.Label(main_frame,
                           text="😂 2025 FJY鎏金PPT生成器",
                           font=("微软雅黑", 8),
                           fg="#555555",
                           bg=self.bg_color)
        copyright.pack(side=tk.BOTTOM, pady=5)
    
    def create_input_row(self, parent, row, label_text, text_var, command):
        """创建统一风格的输入行"""
        frame = tk.Frame(parent, bg=self.bg_color)
        frame.pack(fill=tk.X, pady=5)
        
        tk.Label(frame,
               text=label_text,
               font=("微软雅黑", 10),
               fg=self.fg_color,
               bg=self.bg_color,
               width=15,
               anchor=tk.W).pack(side=tk.LEFT)
        
        entry = tk.Entry(frame,
                       textvariable=text_var,
                       font=("微软雅黑", 10),
                       fg=self.entry_fg,
                       bg=self.entry_bg,
                       insertbackground=self.fg_color,
                       relief=tk.SOLID,
                       bd=1,
                       highlightbackground=self.entry_bd,
                       highlightthickness=1,
                       width=40)
        entry.pack(side=tk.LEFT, padx=5)
        
        btn = ttk.Button(frame,
                        text="浏览...",
                        command=command,
                        style="TButton",
                        width=8)
        btn.pack(side=tk.LEFT)
    
    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)
            if not self.output_folder.get():
                self.output_folder.set(folder)
    
    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)
    
    def generate_output_filename(self):
        """生成自动命名的PPT文件名"""
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        
        prefix = ""
        if self.add_prefix.get():
            prefix = self.prefix_text.get() + "_"
        
        filename = f"{prefix}{timestamp}.pptx"
        return os.path.join(self.output_folder.get(), filename)
    
    def get_image_files(self, folder_path):
        """获取文件夹中的所有图片文件并按名称排序"""
        image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff')
        image_files = []
        
        for file in os.listdir(folder_path):
            if file.lower().endswith(image_extensions):
                image_files.append(os.path.join(folder_path, file))
        
        # 按文件名排序
        image_files.sort()
        return image_files
    
    def get_image_dimensions(self, image_path):
        """获取图片原始尺寸(英寸)"""
        with Image.open(image_path) as img:
            dpi = img.info.get('dpi', (96, 96))  # 默认使用96dpi
            width_inch = img.width / dpi[0]
            height_inch = img.height / dpi[1]
            return width_inch, height_inch
    
    def calculate_symmetrical_layout(self, images_per_page, image_files, start_index):
        """中心对称布局算法"""
        # 使用16:9宽屏比例（单位：英寸）
        page_width = 10.0
        page_height = 5.625
        center_x = page_width / 2
        center_y = page_height / 2
        
        # 获取当前页所有图片的原始尺寸
        image_sizes = []
        for i in range(start_index, min(start_index + images_per_page, len(image_files))):
            try:
                width, height = self.get_image_dimensions(image_files[i])
                image_sizes.append((width, height))
            except:
                image_sizes.append((4, 3))  # 默认比例4:3
        
        # 计算行列数（接近正方形排列）
        cols = math.ceil(math.sqrt(images_per_page))
        rows = math.ceil(images_per_page / cols)
        
        # 计算单元格基本尺寸（包含间距）
        horizontal_gap = 0.15  # 水平间距
        vertical_gap = 0.15    # 垂直间距
        text_height = 0.4 if self.show_filenames.get() else 0
        
        # 计算内容区域总尺寸
        content_width = page_width * 0.9  # 保留10%边距
        content_height = (page_height * 0.9) - (rows * text_height)
        
        cell_width = (content_width - (cols - 1) * horizontal_gap) / cols
        cell_height = (content_height - (rows - 1) * vertical_gap) / rows
        
        # 计算整体内容区域起始位置（居中）
        total_used_width = cols * cell_width + (cols - 1) * horizontal_gap
        total_used_height = rows * cell_height + (rows - 1) * vertical_gap + rows * text_height
        
        start_x = center_x - total_used_width / 2
        start_y = center_y - total_used_height / 2
        
        # 计算每张图片的位置和大小（允许±15%的比例调整）
        positions = []
        for i in range(images_per_page):
            if i >= len(image_sizes):
                break
                
            col = i % cols
            row = i // cols
            
            # 计算单元格位置
            cell_left = start_x + col * (cell_width + horizontal_gap)
            cell_top = start_y + row * (cell_height + vertical_gap + text_height)
            
            # 获取原始尺寸
            orig_width, orig_height = image_sizes[i]
            orig_ratio = orig_width / orig_height
            
            # 计算单元格可用空间
            max_width = cell_width
            max_height = cell_height
            
            # 计算最佳尺寸（允许±15%的比例调整）
            target_ratio = max_width / max_height
            if abs(orig_ratio - target_ratio) / orig_ratio > 0.15:
                # 如果比例差异超过15%，则进行适度调整
                if orig_ratio > target_ratio:
                    # 宽图：保持宽度，调整高度
                    img_width = max_width
                    img_height = img_width / (orig_ratio * 0.85)  # 高度增加不超过15%
                else:
                    # 高图：保持高度，调整宽度
                    img_height = max_height
                    img_width = img_height * (orig_ratio * 1.15)  # 宽度增加不超过15%
            else:
                # 比例差异在可接受范围内，保持原始比例
                img_width = min(max_width, max_height * orig_ratio)
                img_height = min(max_height, max_width / orig_ratio)
            
            # 计算居中位置
            img_left = cell_left + (cell_width - img_width) / 2
            img_top = cell_top + (cell_height - img_height) / 2
            
            positions.append({
                'left': img_left,
                'top': img_top,
                'width': img_width,
                'height': img_height,
                'text_top': img_top + img_height,  # 文字紧贴图片下方
                'text_height': text_height
            })
        
        return {
            'positions': positions,
            'page_width': page_width,
            'page_height': page_height
        }
    
    def convert_images_to_ppt(self):
        """将图片转换为PPT"""
        source_folder = self.source_folder.get()
        output_folder = self.output_folder.get()
        images_per_page = self.images_per_page.get()
        
        if not source_folder or not os.path.isdir(source_folder):
            messagebox.showerror("错误", "请选择有效的图片源文件夹")
            return
        
        if not output_folder:
            messagebox.showerror("错误", "请指定输出文件夹")
            return
        
        try:
            pythoncom.CoInitialize()  # 初始化COM库
            
            # 获取图片文件
            image_files = self.get_image_files(source_folder)
            total_images = len(image_files)
            
            if total_images == 0:
                messagebox.showwarning("警告", "指定的文件夹中没有找到图片文件")
                return
            
            # 生成输出文件名
            output_ppt = self.generate_output_filename()
            
            # 确保输出文件夹存在
            os.makedirs(output_folder, exist_ok=True)
            
            # 创建PPT（16:9宽屏比例）
            prs = Presentation()
            prs.slide_width = Inches(10)      # 宽屏宽度10英寸
            prs.slide_height = Inches(5.625) # 宽屏高度5.625英寸
            
            # 设置进度条
            self.progress['maximum'] = total_images
            self.progress['value'] = 0
            self.root.update()
            
            # 每N张图片创建一个幻灯片
            for i in range(0, total_images, images_per_page):
                # 计算当前页的布局
                layout = self.calculate_symmetrical_layout(images_per_page, image_files, i)
                
                # 添加新幻灯片
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白幻灯片
                
                # 在当前幻灯片上添加图片
                for j in range(min(images_per_page, len(image_files) - i)):
                    pos = layout['positions'][j]
                    
                    # 添加图片（允许适度调整比例）
                    try:
                        slide.shapes.add_picture(
                            image_files[i + j],
                            Inches(pos['left']), Inches(pos['top']), 
                            Inches(pos['width']), Inches(pos['height'])
                        )
                        
                        # 添加图片名称文本
                        if self.show_filenames.get():
                            textbox = slide.shapes.add_textbox(
                                Inches(pos['left']), 
                                Inches(pos['text_top']), 
                                Inches(pos['width']), 
                                Inches(pos['text_height'])
                            )
                            tf = textbox.text_frame
                            tf.word_wrap = True
                            
                            p = tf.paragraphs[0]
                            p.text = os.path.splitext(os.path.basename(image_files[i + j]))[0]
                            p.font.size = Pt(16)  # 16号字体
                            p.font.bold = True    # 加粗
                            p.font.name = "Arial"
                            p.alignment = PP_ALIGN.CENTER
                    
                    except Exception as e:
                        print(f"无法添加图片 {image_files[i + j]}: {str(e)}")
                    
                    # 更新进度
                    self.progress['value'] = i + j + 1
                    self.root.update()
            
            # 保存PPT
            prs.save(output_ppt)
            
            # 完成提示
            messagebox.showinfo(
                "完成", 
                f"转换完成!\n共处理 {total_images} 张图片\n输出文件: {output_ppt}"
            )
            self.progress['value'] = 0
            
        except Exception as e:
            messagebox.showerror("错误", f"转换过程中发生错误:\n{str(e)}")
        finally:
            pythoncom.CoUninitialize()

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToPPTConverter(root)
    root.mainloop()

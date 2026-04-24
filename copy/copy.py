import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


class FileCopyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件夹分层复制工具")
        self.root.geometry("660x460")
        self.root.resizable(False, False)

        self.src_var   = tk.StringVar()
        self.dst_var   = tk.StringVar()
        self.depth_var = tk.IntVar(value=1)

        self._build_ui()

    # ─────────────────────────── UI ───────────────────────────
    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # ① 源文件夹
        frm_src = ttk.LabelFrame(self.root, text="① 选择源文件夹（第一层）")
        frm_src.pack(fill="x", **pad)
        ttk.Entry(frm_src, textvariable=self.src_var, width=62).pack(
            side="left", padx=6, pady=6)
        ttk.Button(frm_src, text="浏览…", command=self._pick_src).pack(
            side="left", padx=4, pady=6)

        # ② 目标文件夹
        frm_dst = ttk.LabelFrame(self.root, text="② 选择目标文件夹")
        frm_dst.pack(fill="x", **pad)
        ttk.Entry(frm_dst, textvariable=self.dst_var, width=62).pack(
            side="left", padx=6, pady=6)
        ttk.Button(frm_dst, text="浏览…", command=self._pick_dst).pack(
            side="left", padx=4, pady=6)

        # ③ 层数
        frm_depth = ttk.LabelFrame(self.root, text="③ 向内复制层数")
        frm_depth.pack(fill="x", **pad)
        ttk.Label(frm_depth,
                  text="层数（1=仅第一层文件，2=含一级子文件夹，以此类推）：").pack(
            side="left", padx=6, pady=6)
        ttk.Spinbox(frm_depth, from_=1, to=20,
                    textvariable=self.depth_var, width=5).pack(
            side="left", pady=6)

        # 预览区
        frm_prev = ttk.LabelFrame(self.root, text="预览（将被复制的项目）")
        frm_prev.pack(fill="both", expand=True, **pad)

        self.preview_text = tk.Text(
            frm_prev, height=9, state="disabled",
            wrap="none", font=("Consolas", 9))
        sb_y = ttk.Scrollbar(frm_prev, orient="vertical",
                              command=self.preview_text.yview)
        sb_x = ttk.Scrollbar(frm_prev, orient="horizontal",
                              command=self.preview_text.xview)
        self.preview_text.configure(
            yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.preview_text.pack(fill="both", expand=True, padx=4, pady=4)

        # 按钮行
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(pady=6)
        ttk.Button(frm_btn, text="预  览",    width=14,
                   command=self._preview).pack(side="left", padx=10)
        ttk.Button(frm_btn, text="开始复制", width=14,
                   command=self._copy).pack(side="left", padx=10)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.status_var,
                  anchor="w", relief="sunken").pack(fill="x", side="bottom")

    # ──────────────────────── 路径选择 ────────────────────────
    def _pick_src(self):
        p = filedialog.askdirectory(title="选择源文件夹")
        if p:
            self.src_var.set(p)

    def _pick_dst(self):
        p = filedialog.askdirectory(title="选择目标文件夹")
        if p:
            self.dst_var.set(p)

    # ──────────────────────── 核心逻辑 ────────────────────────
    def _collect_items(self, src: Path, depth: int):
        """
        返回 (src_path, rel_path) 列表。
        depth=1 → 第一层的文件（不含子文件夹内容）
        depth=2 → 第一层文件 + 第一层子文件夹及其直接内容 …
        """
        items = []

        def _walk(folder: Path, cur: int, rel_base: Path):
            if cur > depth:
                return
            try:
                entries = sorted(folder.iterdir())
            except PermissionError:
                return
            for e in entries:
                rel = rel_base / e.name
                if e.is_file():
                    items.append((e, rel))
                elif e.is_dir():
                    if cur < depth:
                        _walk(e, cur + 1, rel)
                    else:
                        # 已到边界，只记录文件夹节点本身（不展开）
                        items.append((e, rel))

        _walk(src, 1, Path())
        return items

    def _validate_and_collect(self):
        src_s = self.src_var.get().strip()
        dst_s = self.dst_var.get().strip()

        if not src_s:
            messagebox.showwarning("提示", "请先选择源文件夹！"); return None
        if not dst_s:
            messagebox.showwarning("提示", "请先选择目标文件夹！"); return None

        src = Path(src_s)
        dst = Path(dst_s)

        if not src.is_dir():
            messagebox.showerror("错误", f"源文件夹不存在：\n{src}"); return None

        # 防止目标是源的子目录
        try:
            dst.relative_to(src)
            messagebox.showerror("错误", "目标文件夹不能是源文件夹的子目录！")
            return None
        except ValueError:
            pass

        depth = self.depth_var.get()
        items = self._collect_items(src, depth)

        if not items:
            messagebox.showinfo("提示", "在指定层数内未找到任何文件或子文件夹。")
            return None

        return items

    def _preview(self):
        items = self._validate_and_collect()
        if items is None:
            return
        self._show_preview(items)

    def _copy(self):
        items = self._validate_and_collect()
        if items is None:
            return

        src = Path(self.src_var.get().strip())
        dst = Path(self.dst_var.get().strip()) / src.name

        total = len(items)
        if not messagebox.askyesno(
            "确认复制",
            f"共 {total} 个项目\n"
            f"源：{src}\n"
            f"目标：{dst}\n\n"
            "确认开始复制？"
        ):
            return

        copied, skipped, errors = 0, 0, []

        for src_path, rel in items:
            target = dst / rel
            try:
                if src_path.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                else:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src_path, target)
                    copied += 1
            except Exception as e:
                errors.append(f"{rel}: {e}")
                skipped += 1

        msg = f"完成！复制 {copied} 个文件"
        if skipped:
            msg += f"，失败 {skipped} 个"
        self.status_var.set(msg)

        detail = (f"复制完成！\n"
                  f"文件：{copied} 个\n"
                  f"失败：{skipped} 个\n"
                  f"目标路径：{dst}")
        if errors:
            detail += "\n\n错误详情（最多显示 10 条）：\n" + "\n".join(errors[:10])
        messagebox.showinfo("复制结果", detail)

    def _show_preview(self, items):
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        for src_path, rel in items:
            tag = "[文件夹]" if src_path.is_dir() else "[文件  ]"
            self.preview_text.insert("end", f"{tag}  {rel}\n")
        self.preview_text.configure(state="disabled")
        self.status_var.set(f"预览：共 {len(items)} 个项目（点击「开始复制」执行）")


# ───────────────────────── 入口 ─────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    FileCopyApp(root)
    root.mainloop()

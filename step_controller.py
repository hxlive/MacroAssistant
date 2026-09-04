# -*- coding: utf-8 -*-
# step_controller.py
# 功能说明：宏步骤控制器，负责步骤数据、列表展示、参数编辑及手工定位测试
"""统一管理宏步骤状态、步骤配置界面及相关交互逻辑。"""
# Version: 1.8.6

from __future__ import annotations

import copy
import math
import os
import threading
from dataclasses import dataclass
from typing import Any, Callable, Mapping

import pyautogui

from core_engine import MacroSchema
from sys_utils import HotkeyUtils, ImageTooltipManager, RegionPreviewOverlay, RegionSelector
import gui_utils
import screen_locator
from gui_utils import (
    FIND_REGION_MODE_DISPLAY_BY_VALUE,
    FIND_REGION_MODE_OPTIONS,
    param_internal_to_display,
    update_find_region_params,
    update_loop_params,
    update_run_params,
)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class _ParamScrollbarController:
    """自绘浮动滚动条控制器，封装所有滚动条交互细节。

    将原本散落在 MacroApp 中的 18 个方法和 7 个状态变量内聚到一处，
    对外只暴露 frame / canvas / window_id 和 refresh()。
    """

    def __init__(self, parent, root):
        self.root = root
        self.canvas = tk.Canvas(parent, highlightthickness=0, borderwidth=0)
        self.scrollbar = tk.Canvas(self.canvas, width=8, highlightthickness=0, borderwidth=0)
        self.frame = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        # 内部状态
        self._visible = False
        self._update_pending = False
        self._drag_offset = 0
        self._thumb = (0, 0)
        self._thumb_color = "#9aa0a6"
        self._thumb_active_color = "#6f767d"

        # 样式与绑定
        self._apply_theme()
        self.canvas.configure(yscrollcommand=self.on_yview_changed)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.frame.bind("<Configure>", lambda e: self.on_frame_configure())
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.scrollbar.bind("<ButtonPress-1>", self._on_scrollbar_press)
        self.scrollbar.bind("<B1-Motion>", self._on_scrollbar_drag)
        self._bind_mousewheel(self.canvas)
        self._bind_mousewheel(self.frame)
        self.update_visibility()

    def _apply_theme(self):
        try:
            bg = self.canvas.cget("background")
        except Exception:
            bg = "#f5f5f5"
        self.scrollbar.configure(background=bg)

    # --- 事件处理 ---

    def on_frame_configure(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._schedule_update()

    def on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)
        self._schedule_update()

    def on_yview_changed(self, *args):
        # yscrollcommand 回调会传 first, last，这里不需要使用
        self._schedule_update()

    def refresh(self):
        """外部入口：刷新滚动区域、重新绑定子 widget 鼠标滚轮、重置到顶部。"""
        self.root.update_idletasks()
        self._bind_mousewheel_recursive(self.frame)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.yview_moveto(0)
        self.update_visibility()

    # --- 内部更新调度 ---

    def _schedule_update(self):
        if self._update_pending:
            return
        self._update_pending = True
        self.root.after_idle(self.update_visibility)

    def update_visibility(self):
        self._update_pending = False
        bbox = self.canvas.bbox("all")
        content_height = (bbox[3] - bbox[1]) if bbox else 0
        canvas_height = max(self.canvas.winfo_height(), 1)
        should_show = content_height > canvas_height + 1
        self._set_visible(should_show)
        if not should_show:
            self.canvas.yview_moveto(0)
        self._draw_thumb()

    def _set_visible(self, visible):
        if self._visible == visible:
            return
        self._visible = visible
        if visible:
            self.scrollbar.place(in_=self.canvas, relx=1.0, rely=0.0, relheight=1.0, anchor="ne", width=8)
        else:
            self.scrollbar.place_forget()

    def _draw_thumb(self, active=False):
        if not self._visible:
            self.scrollbar.delete("all")
            return
        self.scrollbar.update_idletasks()
        height = max(self.scrollbar.winfo_height(), self.canvas.winfo_height(), 1)
        first, last = self.canvas.yview()
        visible_fraction = max(last - first, 0.05)
        thumb_height = max(int(height * visible_fraction), 24)
        thumb_height = min(thumb_height, height)
        track_height = max(height - thumb_height, 1)
        y1 = int(track_height * first / max(1.0 - visible_fraction, 0.0001)) if visible_fraction < 1 else 0
        y1 = max(0, min(y1, height - thumb_height))
        y2 = y1 + thumb_height
        self._thumb = (y1, y2)
        color = self._thumb_active_color if active else self._thumb_color
        self.scrollbar.delete("all")
        self.scrollbar.create_rectangle(2, y1, 6, y2, fill=color, outline="")

    # --- 拖拽 ---

    def _on_scrollbar_press(self, event):
        if not self._visible:
            return "break"
        y1, y2 = self._thumb
        if y1 <= event.y <= y2:
            self._drag_offset = event.y - y1
        else:
            self._drag_offset = max((y2 - y1) // 2, 0)
            self._move_thumb_to(event.y)
        self._draw_thumb(active=True)
        return "break"

    def _on_scrollbar_drag(self, event):
        if not self._visible:
            return "break"
        self._move_thumb_to(event.y)
        self._draw_thumb(active=True)
        return "break"

    def _move_thumb_to(self, y):
        height = max(self.scrollbar.winfo_height(), 1)
        y1, y2 = self._thumb
        thumb_height = max(y2 - y1, 1)
        track_height = max(height - thumb_height, 1)
        target = max(0, min(y - self._drag_offset, track_height))
        self.canvas.yview_moveto(target / track_height)

    # --- 鼠标滚轮 ---

    def _bind_mousewheel(self, widget):
        if getattr(widget, '_macromate_param_scroll_bound', False):
            return
        widget.bind("<MouseWheel>", self._on_mousewheel, add="+")
        widget.bind("<Button-4>", self._on_mousewheel, add="+")
        widget.bind("<Button-5>", self._on_mousewheel, add="+")
        widget._macromate_param_scroll_bound = True

    def _bind_mousewheel_recursive(self, widget):
        self._bind_mousewheel(widget)
        for child in widget.winfo_children():
            self._bind_mousewheel_recursive(child)

    def _on_mousewheel(self, event):
        if getattr(event, 'num', None) == 4:
            direction = -1
        elif getattr(event, 'num', None) == 5:
            direction = 1
        else:
            direction = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(direction, "units")
        return "break"

def _preview_goto_label(p):
    label = p.get('label', '')
    max_jumps = p.get('max_jumps', 100)
    return f"-> {label}  [最多 {max_jumps} 次]"
def _preview_foreach_line(p):
    source = p.get('file_path') or p.get('source_text', '')
    line_var = p.get('current_line_var', 'current_line')
    fields = p.get('field_names', '')
    suffix = f"；拆分为 {fields}" if fields else ""
    return f"批量处理文本行 '{source}' -> {{{line_var}}}{suffix}"

_STEP_PARAM_PREVIEW_FORMATTERS = {
    'GOTO_LABEL': _preview_goto_label,
    'SET_VAR': lambda p: f"变量 {p.get('var_name', '')} = '{p.get('var_value', '')}'",
    'READ_FILE': lambda p: f"读取文本 '{p.get('file_path', '')}' -> 变量 {p.get('var_name', '')}",
    'EXTRACT_VAR': lambda p: f"'{p.get('source_text', '')}' 提取 '{p.get('regex', '')}' -> 变量 {p.get('var_name', '')}",
    'PROMPT_INPUT': lambda p: f"人工输入 '{p.get('prompt', '')}' -> 变量 {p.get('var_name', '')}",
    'FOREACH_LINE': _preview_foreach_line,
    'END_FOREACH': lambda _p: '结束批量处理',
    'IF_VAR': lambda p: f"如果 '{p.get('var_value', '')}' {p.get('operator', '==')} '{p.get('expected_val', '')}'",
    'CALCULATE': lambda p: f"变量计算 '{p.get('expression', '')}' -> 变量 {p.get('var_name', '')}",
    'WRITE_FILE': lambda p: f"写入文本至 '{p.get('file_path', '')}'",
    'GOTO_IF': lambda p: f"如果 '{p.get('var_value', '')}' {p.get('operator', '==')} '{p.get('expected_val', '')}' -> 跳转至 {p.get('label', '')}",
}
_LIST_DEDENT_ACTIONS = {'ELSE', 'END_IF', 'END_LOOP', 'END_FOREACH'}
_LIST_BLOCK_START_ACTIONS = {'LOOP_START', 'FOREACH_LINE'}
_LIST_BLOCK_END_ACTIONS = {'END_IF', 'END_LOOP', 'END_FOREACH'}


def _is_list_block_start(action):
    return action.startswith('IF_') or action in _LIST_BLOCK_START_ACTIONS


@dataclass(frozen=True)
class StepControllerServices:
    """Application services used by the controller without importing MacroApp."""

    post_to_ui: Callable[[Callable[[], None]], None]
    is_window_alive: Callable[[], bool]
    set_status: Callable[[str], None]
    get_enhanced_mode: Callable[[], bool] = lambda: False
    set_key_recording_active: Callable[[bool], None] = lambda _active: None
    get_reserved_hotkeys: Callable[[], Mapping[str, str]] = lambda: {}
    set_coordinate_capture: Callable[[Callable[[int, int], None] | None], None] = lambda _callback: None


class StepController:
    """Single owner of editable macro-step state and its future widgets."""

    def __init__(
            self,
            root: Any = None,
            widget_factory: Any = None,
            mouse_tracker: Any = None,
            ocr_engine_mapping: Mapping[str, str] | None = None,
            services: StepControllerServices | None = None):
        self.root = root
        self.widget_factory = widget_factory
        self.mouse_tracker = mouse_tracker
        self.font_ui = getattr(widget_factory, 'font_ui', ('Microsoft YaHei UI', 10))
        self.font_code = getattr(widget_factory, 'font_code', ('Consolas', 10))
        self.ocr_engine_mapping = dict(ocr_engine_mapping or {})
        self.ocr_key_mapping = {name: key for key, name in self.ocr_engine_mapping.items()}
        self.services = services or StepControllerServices(
            post_to_ui=lambda callback: callback(),
            is_window_alive=lambda: True,
            set_status=lambda _text: None,
        )

        self.steps: list[dict[str, Any]] = []
        self.editing_index: int | None = None
        self.last_test_location: tuple[int, int] | None = None
        self.last_test_locate_anchor: tuple[int, int] | None = None
        self.editing_step_has_cache_box = False
        self.available_ocr_keys = list(self.ocr_engine_mapping)

        # Widget ownership is declared now; widgets are created only by build methods.
        self.steps_tree = None
        self.tree_menu = None
        self.param_widgets: dict[str, Any] = {}
        self.param_frame = None
        self.action_type = None
        self._clear_selection_after_id = None
        self.add_step_btn = None
        self.cancel_edit_btn = None
        self.remove_btn = None
        self.move_up_btn = None
        self.move_down_btn = None
        self.load_step_btn = None
        self.tooltip_manager = None
        self.scroll_controller = None
        self._param_scroll_ctrl = None
        self.mouse_pos_var = getattr(mouse_tracker, 'var', None)


        self._coordinate_capture_generation = 0

    def build_step_tree(self, list_frame):
        title_frame = ttk.Frame(list_frame)
        title_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(title_frame, text="宏步骤序列:", font=("Microsoft YaHei UI", 11, "bold")).pack(side=tk.LEFT)

        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("id", "action", "params")
        self.steps_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

        self.steps_tree.heading("id", text="#")
        self.steps_tree.heading("action", text="动作")
        self.steps_tree.heading("params", text="参数详情 / 备注")

        self.steps_tree.column("id", width=45, minwidth=40, stretch=False, anchor="center")
        self.steps_tree.column("action", width=220, minwidth=200, stretch=False)
        self.steps_tree.column("params", width=320, minwidth=280, stretch=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.steps_tree.yview)
        self.steps_tree.configure(yscrollcommand=scrollbar.set)

        self.steps_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.steps_tree.bind("<Double-1>", lambda e: self.load_step_for_edit())

        self.tree_menu = tk.Menu(self.root, tearoff=0, font=self.font_ui)
        self.tree_menu.add_command(label="屏蔽/启用选中步骤", command=self.toggle_step_enabled)
        self.steps_tree.bind("<Button-3>", self.show_tree_menu)

        self.steps_tree.tag_configure('editing', background='#FFF3CD')
        self.steps_tree.tag_configure('disabled', foreground='#999999')

    def _get_selected_index(self):
        """获取当前选中项的索引"""
        selected_items = self.steps_tree.selection()
        if not selected_items: return None
        return self.steps_tree.index(selected_items[0])

    def _format_step_params(self, step, act):
        # 参数预览文本
        display_params = self._params_for_edit(act, step['params'])

        cache_str = ""
        if 'region' in display_params or 'cache_box' in display_params:
            box = display_params.pop('region', display_params.pop('cache_box', None))
            if isinstance(box, (list, tuple)) and len(box) >= 4:
                cache_str = f"[区域: {box[0]},{box[1]},{box[2]},{box[3]}] "
            elif box is not None:
                cache_str = "[区域: 无效] "

        if 'engine' in display_params:
            # <--- 列表显示时也使用完整映射
            display_params['engine'] = getattr(self, 'ocr_engine_mapping', getattr(self, 'FULL_OCR_NAME_MAP', {})).get(display_params['engine'], display_params['engine'])

        # 格式化参数列字符串
        param_text = f"{cache_str}{display_params}" if display_params else ""

        # 备注动作特殊处理：显示为注释格式
        if act == 'NOTE':
            note_text = step['params'].get('text', '')
            param_text = f"// {note_text}" if note_text else "// (空备注)"

        formatter = _STEP_PARAM_PREVIEW_FORMATTERS.get(act)
        if formatter:
            param_text = formatter(step['params'])

        # 插入行 (Values对应: id, action, params)
        return param_text

    def _get_step_display_indent(self, action, block_stack):
        return max(0, len(block_stack) - (1 if action in _LIST_DEDENT_ACTIONS else 0))

    def _update_display_block_stack(self, action, block_stack):
        if _is_list_block_start(action):
            block_stack.append(action)
        elif action in _LIST_BLOCK_END_ACTIONS and block_stack:
            block_stack.pop()

    def _get_step_tree_tags(self, index, is_enabled):
        tags = []
        if index == self.editing_index:
            tags.append('editing')
        if not is_enabled:
            tags.append('disabled')
        return tuple(tags)

    def _build_step_tree_row(self, index, step, block_stack):
        act = step['action']
        indent_str = "    " * self._get_step_display_indent(act, block_stack)
        param_text = self._format_step_params(step, act)
        action_label = MacroSchema.ACTION_TRANSLATIONS.get(act, act)
        is_enabled = step.get('enabled', True)

        display_action = f"{indent_str}{action_label}"
        if not is_enabled:
            display_action = f"{indent_str}[屏蔽] {action_label}"

        values = (index + 1, display_action, param_text)
        tags = self._get_step_tree_tags(index, is_enabled)
        return act, values, tags

    def _select_step_tree_item(self, item_id, ensure_visible=True):
        if ensure_visible:
            self.steps_tree.see(item_id)
        self.steps_tree.selection_set(item_id)

    def _select_step_tree_index(self, index, ensure_visible=True):
        children = self.steps_tree.get_children()
        if 0 <= index < len(children):
            self._select_step_tree_item(children[index], ensure_visible)
            return True
        return False

    def _focus_step_tree_item_if_editing(self, index, item_id):
        if index == self.editing_index:
            self._select_step_tree_item(item_id)

    def update_listbox_display(self):
        """Refresh the Treeview display."""
        for item in self.steps_tree.get_children():
            self.steps_tree.delete(item)

        block_stack = []
        for i, step in enumerate(self.steps):
            act, values, tags = self._build_step_tree_row(i, step, block_stack)
            item_id = self.steps_tree.insert("", "end", values=values)
            if tags:
                self.steps_tree.item(item_id, tags=tags)

            self._focus_step_tree_item_if_editing(i, item_id)
            self._update_display_block_stack(act, block_stack)



    def _is_step_toggle_allowed(self, index):
        action = self.steps[index].get('action', '')
        return action not in MacroSchema.CONTROL_FLOW_ACTIONS

    def _set_tree_menu_toggle_state(self, index):
        state = "normal" if self._is_step_toggle_allowed(index) else "disabled"
        self.tree_menu.entryconfig(0, state=state)

    def _get_context_menu_step_index(self, event):
        item = self.steps_tree.identify_row(event.y)
        if not item:
            return None

        self._select_step_tree_item(item, ensure_visible=False)
        return self._get_selected_index()

    def _warn_control_flow_toggle_blocked(self):
        messagebox.showwarning("提示", "不可屏蔽流程控制节点（条件、循环），以防止引发严重 BUG。", parent=self.root)

    def _toggle_step_enabled_at(self, index):
        step = self.steps[index]
        step['enabled'] = not step.get('enabled', True)

    def show_tree_menu(self, event):
        """Show the step list context menu."""
        idx = self._get_context_menu_step_index(event)
        if idx is None:
            return

        self._set_tree_menu_toggle_state(idx)
        self.tree_menu.post(event.x_root, event.y_root)

    def toggle_step_enabled(self):
        """切换选中步骤的启用/屏蔽状态"""
        idx = self._get_selected_index()
        if idx is None:
            return

        if not self._is_step_toggle_allowed(idx):
            self._warn_control_flow_toggle_blocked()
            return

        self._toggle_step_enabled_at(idx)
        self.update_listbox_display()

    def remove_step(self):
        # --- 升级: 适配 Treeview ---
        idx = self._get_selected_index()
        if idx is None: return

        # [修复] 使用 elif 确保逻辑互斥
        # 原代码问题: cancel_edit_mode 会将 editing_index 设为 None，
        # 导致后续的 if 判断永远为 False，索引调整失效
        if self.editing_index == idx:
            self.cancel_edit_mode()
        elif self.editing_index is not None and self.editing_index > idx:
            self.editing_index -= 1

        del self.steps[idx]
        self.update_listbox_display()

        # 尝试选中下一行
        children = self.steps_tree.get_children()
        if idx < len(children):
             self._select_step_tree_index(idx, ensure_visible=False)
        elif children:
             self._select_step_tree_index(len(children) - 1, ensure_visible=False)

    def move_step(self, d):
        # --- 升级: 适配 Treeview ---
        idx = self._get_selected_index()
        if idx is None: return

        i = idx
        new_i = i - 1 if d == "up" else i + 1

        if 0 <= new_i < len(self.steps):
            self.steps.insert(new_i, self.steps.pop(i))

            # 同步更新 editing_index
            if self.editing_index == i: self.editing_index = new_i
            elif self.editing_index == new_i: self.editing_index = i
            self.update_listbox_display()

            # 保持选中移动后的项
            self._select_step_tree_index(new_i)


    def build_step_form(self, main_frame):
        add_frame = ttk.Labelframe(main_frame, text="添加新步骤", padding=10)
        add_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)

        add_frame.pack_propagate(False)
        add_frame.configure(width=380)

        right_bottom_frame = ttk.Frame(add_frame)
        right_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        right_bottom_frame.columnconfigure(0, weight=2); right_bottom_frame.columnconfigure(1, weight=1)

        self.add_step_btn = ttk.Button(right_bottom_frame, text="＋ 添加到序列 >>", command=self.add_or_update_step, bootstyle="success", padding=(12, 8))
        self.add_step_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2), columnspan=2)
        self.cancel_edit_btn = ttk.Button(right_bottom_frame, text="✕ 取消修改", command=self.cancel_edit_mode, bootstyle="secondary", padding=(10, 6))

        ttk.Label(add_frame, text="选择动作:").pack(anchor="w")
        self.action_type = ttk.Combobox(add_frame, state="readonly", font=self.font_ui, height=16)
        self.action_type['values'] = list(MacroSchema.ACTION_TRANSLATIONS.values())
        self.action_type.current(0)

        self.action_type.pack(anchor="w", fill=tk.X, pady=5)
        self.action_type.bind("<<ComboboxSelected>>", self._on_action_type_selected)
        self.action_type.bind("<FocusIn>", self._clear_combobox_text_selection, add="+")
        self.action_type.bind("<ButtonRelease-1>", self._clear_combobox_text_selection, add="+")
        self.action_type.bind("<Destroy>", self._cancel_clear_combobox_callback, add="+")

        param_area = ttk.Frame(add_frame)
        param_area.pack(fill=tk.BOTH, expand=True, pady=5)
        self.build_scrollable_param_area(param_area)

        self.param_widgets = {}

    def _clear_combobox_text_selection(self, event=None):
        widget = getattr(event, 'widget', None) if event is not None else self.action_type
        if widget is None:
            return

        self._cancel_clear_combobox_callback(event)

        def clear_selection():
            self._clear_selection_after_id = None
            try:
                widget.selection_clear()
                widget.icursor(tk.END)
            except tk.TclError:
                pass

        try:
            self._clear_selection_after_id = widget.after_idle(clear_selection)
        except tk.TclError:
            pass

    def _cancel_clear_combobox_callback(self, event=None):
        after_id = self._clear_selection_after_id
        self._clear_selection_after_id = None
        if after_id is None:
            return

        widget = getattr(event, 'widget', None) if event is not None else self.action_type
        if widget is None:
            return
        try:
            widget.after_cancel(after_id)
        except tk.TclError:
            pass

    def _on_action_type_selected(self, event):
        self.update_param_fields(event)
        self._clear_combobox_text_selection(event)

    def _apply_loop_mode_for_edit(self, params):
        saved_mode = params.get('mode', 'fixed')
        default_display = next(iter(gui_utils.LOOP_MODE_OPTIONS.keys()))
        display_mode = gui_utils.LOOP_MODE_DISPLAY_BY_VALUE.get(saved_mode, default_display)
        mode_widget = self.param_widgets.get('mode')
        if mode_widget is not None:
            mode_widget.set(display_mode)
            update_loop_params(self.param_widgets, self.param_frame, mode_widget)

    def _apply_run_type_for_edit(self, params):
        saved_run_type = params.get('run_type', 'command')
        default_display = next(iter(MacroSchema.RUN_TYPE_OPTIONS.keys()))
        display_run_type = MacroSchema.RUN_TYPE_DISPLAY_BY_VALUE.get(saved_run_type, default_display)
        run_type_widget = self.param_widgets.get('run_type')
        if run_type_widget is not None:
            run_type_widget.set(display_run_type)
            update_run_params(self.param_widgets, self.param_frame, run_type_widget)

    def _apply_find_region_mode_for_edit(self, params):
        mode_widget = self.param_widgets.get('region_mode')
        if mode_widget is None:
            return
        saved_mode = params.get('region_mode')
        if saved_mode is None:
            saved_mode = (
                'absolute'
                if params.get('region') is not None or params.get('cache_box') is not None
                else 'full'
            )
        display_mode = FIND_REGION_MODE_DISPLAY_BY_VALUE.get(saved_mode, saved_mode)
        mode_widget.set(display_mode)
        update_find_region_params(
            self.param_widgets, self.param_frame, mode_widget
        )

    def _prepare_action_form_for_edit(self, step):
        self.action_type.set(MacroSchema.ACTION_TRANSLATIONS.get(step['action']))
        self.update_param_fields(None)
        self._clear_combobox_text_selection()

        if step['action'] == 'LOOP_START':
            self._apply_loop_mode_for_edit(step['params'])
        elif step['action'] == 'RUN':
            self._apply_run_type_for_edit(step['params'])
        elif step['action'] in (
                'FIND_IMAGE', 'IF_IMAGE_FOUND', 'FIND_TEXT', 'IF_TEXT_FOUND'):
            self._apply_find_region_mode_for_edit(step['params'])

    def _fill_region_param_for_edit(self, params):
        if 'region' not in self.param_widgets:
            return
        region_box = params.get('region', params.get('cache_box'))
        self.editing_step_has_cache_box = ('cache_box' in params and 'region' not in params)
        if isinstance(region_box, list) and len(region_box) == 4:
            self._set_entry_text(
                self.param_widgets['region'],
                self._format_region_box(region_box),
            )

    def _display_value_for_param(self, key, value):
        if key == 'comparison':
            return gui_utils.COLOR_COMPARISON_DISPLAY_BY_VALUE.get(value, value)
        if key not in ('lang', 'button', 'engine'):
            return value
        return param_internal_to_display(
            key, value,
            self.ocr_engine_mapping,
            MacroSchema.LANG_VALUES_TO_NAME,
            MacroSchema.CLICK_VALUES_TO_NAME,
            self.available_ocr_keys
        )

    def _fill_regular_params_for_edit(self, params):
        for key, value in params.items():
            if key in ('mode', 'run_type', 'region_mode', 'cache_box', 'region'):
                continue
            if key not in self.param_widgets:
                continue
            self._set_param_widget_value(key, self._display_value_for_param(key, value))

    @staticmethod
    def _params_for_edit(action, params):
        """Return form values with legacy action times converted to milliseconds."""
        converted = copy.deepcopy(params)
        specs = {
            'CLICK': {
                'interval_ms': ('interval', 1),
                'duration_ms': ('duration', 1000),
            },
            'MOVE_TO': {'duration_ms': ('duration', 1000)},
            'MOVE_OFFSET': {'duration_ms': ('duration', 1000)},
            'TYPE_TEXT': {'interval_ms': ('interval', 1000)},
            'AI_COMMAND': {'duration_ms': ('duration', 1000)},
            'FIND_IMAGE': {'retry_interval_ms': ('retry_interval', 1000)},
            'IF_IMAGE_FOUND': {'retry_interval_ms': ('retry_interval', 1000)},
            'FIND_TEXT': {'retry_interval_ms': ('retry_interval', 1000)},
            'IF_TEXT_FOUND': {'retry_interval_ms': ('retry_interval', 1000)},
            'RUN': {'timeout_ms': ('timeout', 1000)},
        }
        for new_key, (legacy_key, multiplier) in specs.get(action, {}).items():
            if new_key in converted:
                converted.pop(legacy_key, None)
                continue
            if legacy_key not in converted:
                continue
            try:
                value = float(converted[legacy_key]) * multiplier
            except (TypeError, ValueError, OverflowError):
                continue
            converted[new_key] = int(value) if value.is_integer() else value
            converted.pop(legacy_key, None)
        return converted

    def _enter_edit_mode(self, index):
        self.editing_index = index
        self.add_step_btn.config(text="[OK] 更新步骤", bootstyle="warning")
        self.add_step_btn.grid_configure(columnspan=1)
        self.cancel_edit_btn.grid(row=0, column=1, sticky="nsew", padx=(2, 0))
        self.update_listbox_display()

    def load_step_for_edit(self):
        """加载选中步骤到编辑区。"""
        idx = self._get_selected_index()
        if idx is None:
            return

        step = self.steps[idx]
        self._prepare_action_form_for_edit(step)
        self._fill_region_param_for_edit(step['params'])
        self._fill_regular_params_for_edit(
            self._params_for_edit(step['action'], step['params'])
        )
        self._enter_edit_mode(idx)

    def cancel_edit_mode(self):
        self.editing_index = None
        self.editing_step_has_cache_box = False
        if self.add_step_btn is not None:
            self.add_step_btn.config(text="＋ 添加到序列 >>", bootstyle="success")
            self.add_step_btn.grid_configure(columnspan=2)
        if self.cancel_edit_btn is not None:
            self.cancel_edit_btn.grid_remove()
        if self.steps_tree is not None:
            self.update_listbox_display()


    def _parse_optional_test_region(self, region_value):
        raw_region = (region_value or '').strip()
        if not raw_region:
            return None

        region_box = gui_utils.parse_region_string(raw_region)
        if region_box is None or screen_locator.bbox_to_region(region_box) is None:
            raise ValueError('手动查找区域格式无效，应为 x1, y1, x2, y2')
        return region_box

    def on_preview_region(self, entry_widget):
        raw_region = entry_widget.get().strip()
        if not raw_region:
            messagebox.showinfo('区域预览', '请先框选或输入搜索范围。')
            return

        try:
            region = self._parse_optional_test_region(raw_region)
            RegionPreviewOverlay(self.root, region, duration_ms=1500)
        except (TypeError, ValueError, tk.TclError) as exc:
            messagebox.showwarning('区域预览', f'无法显示搜索范围：{exc}')

    def on_select_region(self, entry_widget):
        self.root.iconify()
        self.root.after(300, lambda: self._do_select_region(entry_widget))

    def _do_select_region(self, entry_widget):
        try:
            region = RegionSelector(self.root).get_region()
            self.root.deiconify()
            if region:
                self._set_entry_text(entry_widget, self._format_region_box(region))
        except Exception as exc:
            self.root.deiconify()
            messagebox.showerror('错误', f'选区失败: {exc}')

    def _read_test_region_box(self):
        mode_widget = self.param_widgets.get('region_mode')
        if mode_widget is not None:
            display_mode = mode_widget.get()
            mode = FIND_REGION_MODE_OPTIONS.get(display_mode, display_mode)
            if mode == 'full':
                return None
            if mode == 'relative':
                if self.last_test_locate_anchor is None:
                    raise ValueError('相对搜索范围没有参考点，请先测试一次锚点识别步骤')
                try:
                    x_offset = int(self.param_widgets['region_x_offset'].get())
                    y_offset = int(self.param_widgets['region_y_offset'].get())
                    width = int(self.param_widgets['region_width'].get())
                    height = int(self.param_widgets['region_height'].get())
                except (KeyError, TypeError, ValueError, OverflowError) as exc:
                    raise ValueError('相对搜索范围的偏移、宽度和高度必须是整数') from exc
                if width <= 0 or height <= 0:
                    raise ValueError('相对搜索范围的宽度和高度必须大于 0')
                anchor_x, anchor_y = self.last_test_locate_anchor
                left = anchor_x + x_offset
                top = anchor_y + y_offset
                return (left, top, left + width, top + height)
            if mode != 'absolute':
                raise ValueError(f'不支持的搜索范围模式: {display_mode}')

        region_widget = self.param_widgets.get('region')
        if region_widget is None:
            return None
        return self._parse_optional_test_region(region_widget.get())

    def on_test_find_image_click(self):
        try:
            path = self.param_widgets['path'].get()
            confidence = float(self.param_widgets['confidence'].get())
            if not math.isfinite(confidence) or not 0.0 <= confidence <= 1.0:
                raise ValueError('置信度必须是 0 到 1 之间的有限数字')
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            request = screen_locator.LocateRequest(
                mode='image',
                region_bbox=self._read_test_region_box(),
                template_path=path,
                confidence=confidence,
                enhanced_mode=bool(self.services.get_enhanced_mode()),
            )
            self._begin_background_test('测试中...', request)
        except Exception as exc:
            messagebox.showerror('错误', f'参数无效: {exc}')

    def on_test_find_text_click(self):
        try:
            text = self.param_widgets['text'].get()
            lang = MacroSchema.LANG_OPTIONS.get(
                self.param_widgets['lang'].get(), 'eng'
            )
            engine_name = self.param_widgets['engine'].get()
            if not text:
                raise ValueError('查找文本不能为空')
            if engine_name.endswith(' (不可用)'):
                messagebox.showwarning(
                    '引擎不可用',
                    f"您选择的引擎 '{engine_name}' 在当前环境中未安装或无法加载。\n\n"
                    '请选择其他引擎，或安装相应组件后重启程序。',
                    parent=self.root,
                )
                return
            request = screen_locator.LocateRequest(
                mode='text',
                region_bbox=self._read_test_region_box(),
                target_text=text,
                lang=lang,
                ocr_engine=self.ocr_key_mapping.get(engine_name, 'auto'),
                ocr_debug=True,
                enhanced_mode=bool(self.services.get_enhanced_mode()),
            )
            self._begin_background_test('测试中...', request)
        except Exception as exc:
            messagebox.showerror('错误', f'参数无效: {exc}')

    def on_test_ai_command_click(self):
        try:
            instruction = self.param_widgets['instruction'].get()
            if not instruction.strip():
                messagebox.showwarning('提示', '请输入 AI 指令')
                return
            request = screen_locator.LocateRequest(
                mode='ai',
                region_bbox=self._read_test_region_box(),
                instruction=instruction,
                enhanced_mode=bool(self.services.get_enhanced_mode()),
            )
            self._begin_background_test('AI 分析中...', request)
        except Exception as exc:
            messagebox.showerror('错误', f'参数无效: {exc}')

    def _begin_background_test(self, status_text, request):
        self.services.set_status(status_text)
        self.root.iconify()
        self._run_test_after_iconify(request)

    def _run_test_after_iconify(self, request, attempts=0):
        if self.root.state() == 'iconic' or attempts >= 15:
            self.root.after(250, lambda: self._run_test_thread(request))
            return
        self.root.after(
            100, lambda: self._run_test_after_iconify(request, attempts + 1)
        )

    def _run_test_thread(self, request):
        threading.Thread(
            target=self._run_manual_location, args=(request,), daemon=True
        ).start()

    def _run_manual_location(self, request):
        try:
            result = screen_locator.locate(request)
        except Exception as exc:
            self.services.post_to_ui(
                lambda error=exc: self._deliver_test_error(error)
            )
        else:
            self.services.post_to_ui(
                lambda located=result: self._deliver_test_result(located)
            )

    def _deliver_test_result(self, result):
        if not self.services.is_window_alive():
            return
        self._on_test_complete(result)

    def _deliver_test_error(self, error):
        if not self.services.is_window_alive():
            return
        self._on_test_error(error)

    def _restore_main_window_after_test(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)

    def _on_test_complete(self, result):
        self._restore_main_window_after_test()
        try:
            if result.found:
                self.last_test_location = result.position
                if result.source in ('image', 'ocr'):
                    self.last_test_locate_anchor = result.position
                pyautogui.moveTo(*result.position)
                if result.source == 'vlm':
                    messagebox.showinfo(
                        'AI 成功',
                        f'找到坐标: {self.last_test_location}\n\nAI 已移动鼠标到该位置',
                    )
                else:
                    messagebox.showinfo(
                        '成功', f'找到于 {self.last_test_location}'
                    )
            elif result.source == 'vlm':
                messagebox.showwarning(
                    'AI 失败',
                    '未能从 AI 获取有效坐标\n\n请检查:\n'
                    '1. API Key 是否正确配置\n2. 网络是否正常\n3. 指令是否清晰',
                )
            else:
                messagebox.showwarning('失败', '未找到目标')
        finally:
            self.services.set_status('')
            self.root.attributes('-topmost', False)

    def _on_test_error(self, error):
        self._restore_main_window_after_test()
        try:
            messagebox.showerror('错误', str(error))
        finally:
            self.services.set_status('')
            self.root.attributes('-topmost', False)

    def browse_image(self):
        """浏览图片文件（保持向后兼容）"""
        f = filedialog.askopenfilename(filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if f:
            f = os.path.abspath(f)
            self._set_entry_text(self.param_widgets['path'], f)

    def _validate_recorded_key(self, key_chord):
        normalized = HotkeyUtils.normalize_key_chord(key_chord)
        for purpose, configured in self.services.get_reserved_hotkeys().items():
            if normalized == HotkeyUtils.normalize_key_chord(configured):
                display = HotkeyUtils.format_hotkey_display(normalized)
                return (
                    f"无法录入 {display}："
                    f"该组合键已被设置为“{purpose}”快捷键"
                )
        return None

    def _on_key_recording_change(self, active):
        self.services.set_key_recording_active(bool(active))

    def _on_key_recording_error(self, message):
        self.services.set_status(str(message))

    def stop_key_recording(self):
        """Release both the active PRESS_KEY widget and global suppression flag."""
        key_widget = self.param_widgets.get('key')
        stop_recording = getattr(key_widget, 'stop_recording', None)
        if callable(stop_recording):
            try:
                stop_recording()
            except tk.TclError:
                pass
        self.services.set_key_recording_active(False)

    def stop_coordinate_capture(self):
        """Invalidate queued captures and release the temporary F8 target."""
        self._coordinate_capture_generation += 1
        self.services.set_coordinate_capture(None)

    def stop_input_capture(self):
        """Stop every temporary input mode before run/settings/form transitions."""
        self.stop_key_recording()
        self.stop_coordinate_capture()

    @staticmethod
    def _coordinate_entry_is_alive(entry):
        try:
            return not hasattr(entry, 'winfo_exists') or bool(entry.winfo_exists())
        except tk.TclError:
            return False

    def _start_coordinate_capture(self, action, *, show_conflict=True):
        """Route plain F8 samples into the active absolute-coordinate form."""
        self.stop_coordinate_capture()
        if action not in {'MOVE_TO', 'CLICK', 'DRAG_TO', 'IF_COLOR_MATCH'}:
            return

        for purpose, configured in self.services.get_reserved_hotkeys().items():
            if HotkeyUtils.normalize_key_chord(configured) == 'f8':
                if show_conflict:
                    messagebox.showwarning(
                        "F8 快捷键冲突",
                        f"无法启用 F8 坐标取点：F8 已被设置为“{purpose}”快捷键。",
                        parent=self.root,
                    )
                return

        self._coordinate_capture_generation += 1
        generation = self._coordinate_capture_generation
        drag_target = 'start'

        def apply_coordinate(x, y):
            nonlocal drag_target
            if generation != self._coordinate_capture_generation:
                return
            if action == 'DRAG_TO':
                x_key = f'{drag_target}_x'
                y_key = f'{drag_target}_y'
            else:
                x_key, y_key = 'x', 'y'
            x_entry = self.param_widgets.get(x_key)
            y_entry = self.param_widgets.get(y_key)
            if x_entry is None or y_entry is None:
                return
            if not (
                    self._coordinate_entry_is_alive(x_entry)
                    and self._coordinate_entry_is_alive(y_entry)):
                return
            self._set_entry_text(x_entry, int(x))
            self._set_entry_text(y_entry, int(y))
            if action == 'DRAG_TO':
                recorded_target = '起点' if drag_target == 'start' else '终点'
                drag_target = 'end' if drag_target == 'start' else 'start'
                next_target = '起点' if drag_target == 'start' else '终点'
                self.services.set_status(
                    f"F8 已记录拖动{recorded_target}: ({int(x)}, {int(y)})；下一次记录{next_target}"
                )
            elif action == 'IF_COLOR_MATCH':
                try:
                    red, green, blue = screen_locator.sample_screen_pixel(int(x), int(y))
                    color_text = f'#{red:02X}{green:02X}{blue:02X}'
                    color_entry = self.param_widgets.get('target_color')
                    if color_entry is not None and self._coordinate_entry_is_alive(color_entry):
                        self._set_entry_text(color_entry, color_text)
                    self.services.set_status(
                        f"F8 已记录坐标: ({int(x)}, {int(y)})，颜色: {color_text}"
                    )
                except Exception as exc:
                    self.services.set_status(
                        f"F8 已记录坐标: ({int(x)}, {int(y)})，但读取颜色失败: {exc}"
                    )
            else:
                self.services.set_status(
                    f"F8 已记录坐标: ({int(x)}, {int(y)})"
                )

        self.services.set_coordinate_capture(apply_coordinate)
        if action == 'DRAG_TO':
            self.services.set_status(
                "F8 拖动取点已启用：下一次记录起点"
            )
        elif action == 'IF_COLOR_MATCH':
            self.services.set_status(
                "F8 颜色取点已启用：将鼠标移到目标像素后按 F8"
            )
        else:
            self.services.set_status(
                "F8 坐标取点已启用：将鼠标移到目标位置后按 F8"
            )

    def refresh_coordinate_capture(self):
        """Restore F8 capture for the current form after a temporary pause."""
        action = self._get_current_action_key() if self.action_type is not None else None
        self._start_coordinate_capture(action, show_conflict=False)


    def _build_step_from_form(self, action):
        params, error = self.widget_factory.collect_step_data(action, self.param_widgets, self.ocr_key_mapping)
        if error:
            messagebox.showwarning("输入错误", error)
            return None

        if self.editing_step_has_cache_box and 'region' in params:
            params['cache_box'] = params.pop('region')
            params.pop('region_mode', None)


        if action == 'PRESS_KEY':
            conflict = self._validate_recorded_key(params.get('key', ''))
            if conflict:
                messagebox.showwarning("输入错误", conflict)
                return None
        return {"action": action, "params": params}

    def _maybe_apply_test_cache(self, step, action):
        if action not in ('FIND_TEXT', 'FIND_IMAGE', 'IF_TEXT_FOUND', 'IF_IMAGE_FOUND'):
            return
        if self.editing_index is not None or not self.last_test_location:
            return
        if 'region_mode' in step['params']:
            return
        if 'region' in step['params'] or 'cache_box' in step['params']:
            return
        if messagebox.askyesno("缓存", "使用测试坐标作为缓存？"):
            x, y = self.last_test_location
            step["params"]["cache_box"] = [x, y, x + 1, y + 1]

    def _upsert_step(self, step):
        if self.editing_index is not None:
            target_index = self.editing_index
            self.steps[target_index] = step
            self.cancel_edit_mode()
            return target_index

        selected_idx = self._get_selected_index()
        if selected_idx is None:
            self.steps.append(step)
            target_index = len(self.steps) - 1
        else:
            target_index = selected_idx + 1
            self.steps.insert(target_index, step)

        self.update_listbox_display()
        return target_index

    def add_or_update_step(self):
        """添加新步骤或更新当前编辑中的步骤。"""
        action = self._get_current_action_key()
        if not action:
            return

        step = self._build_step_from_form(action)
        if step is None:
            return

        self._maybe_apply_test_cache(step, action)
        target_index = self._upsert_step(step)
        self._select_step_tree_index(target_index)
        self.last_test_location = None


    def _get_current_action_key(self):
        return MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())

    def _set_entry_text(self, entry_widget, value):
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, str(value))

    def _set_param_widget_value(self, key, display_val):
        """按 widget 类型分发赋值（BooleanVar / Combobox / Entry）。"""
        w = self.param_widgets[key]
        if isinstance(w, tk.BooleanVar):
            w.set(display_val)
        elif isinstance(w, ttk.Combobox):
            w.set(display_val)
        else:
            self._set_entry_text(w, display_val)

    def _format_region_box(self, region_box):
        return f"{region_box[0]}, {region_box[1]}, {region_box[2]}, {region_box[3]}"


    def update_param_fields(self, event=None):
        self.last_test_location = None

        # [变更] 停止鼠标追踪
        self.mouse_tracker.stop()
        self.mouse_pos_var.set("")

        self.stop_input_capture()
        for widget in self.param_frame.winfo_children(): widget.destroy()
        self.param_widgets = {}
        action_key = MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action_key: return

        # 准备表单回调函数字典
        callbacks = {
            'on_select_region': self.on_select_region,
            'on_preview_region': self.on_preview_region,
            'browse_image': self.browse_image,
            'on_test_find_image_click': self.on_test_find_image_click,
            'on_test_find_text_click': self.on_test_find_text_click,
            'on_test_ai_command_click': self.on_test_ai_command_click,
            'update_loop_params': lambda event: update_loop_params(self.param_widgets, self.param_frame, self.param_widgets.get('mode')) if self.param_widgets.get('mode') is not None else None,
            'update_run_params': lambda event: update_run_params(self.param_widgets, self.param_frame, self.param_widgets.get('run_type')) if self.param_widgets.get('run_type') is not None else None,
            'mouse_tracker': self.mouse_tracker,
            'mouse_pos_var': self.mouse_pos_var,
            'on_key_recording_change': self._on_key_recording_change,
            'validate_recorded_key': self._validate_recorded_key,
            'on_key_recording_error': self._on_key_recording_error,
        }

        # 使用表单工厂构建参数控件
        res = self.widget_factory.build_action_form(
            action_key,
            self.param_frame,
            self.param_widgets,
            self.available_ocr_keys,
            callbacks
        )

        # 处理特殊返回 (如 OCR 不可用时自动切回图像模式)
        if res == "SWITCH_TO_FIND_IMAGE":
            self.action_type.set(MacroSchema.ACTION_TRANSLATIONS['FIND_IMAGE'])
            self.update_param_fields(None)
            return

        self._start_coordinate_capture(action_key)
        self._param_scroll_ctrl.refresh()

    def build_scrollable_param_area(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        self._param_scroll_ctrl = _ParamScrollbarController(parent, self.root)
        self.scroll_controller = self._param_scroll_ctrl
        self.param_frame = self._param_scroll_ctrl.frame
        return self.param_frame

    def build_step_controls(self, parent, row=0):
        self.move_up_btn = ttk.Button(parent, text="↑ 上移", command=lambda: self.move_step("up"), bootstyle="primary-outline", padding=(10, 6))
        self.move_up_btn.grid(row=row, column=0, sticky="nsew", padx=(0, 2), pady=(0, 5))
        self.move_down_btn = ttk.Button(parent, text="↓ 下移", command=lambda: self.move_step("down"), bootstyle="primary-outline", padding=(10, 6))
        self.move_down_btn.grid(row=row, column=1, sticky="nsew", padx=2, pady=(0, 5))
        self.remove_btn = ttk.Button(parent, text="🗑 删除选中", command=self.remove_step, bootstyle="danger-outline", padding=(10, 6))
        self.remove_btn.grid(row=row, column=2, sticky="nsew", padx=2, pady=(0, 5))
        self.load_step_btn = ttk.Button(parent, text="✎ 修改步骤", command=self.load_step_for_edit, bootstyle="info-outline", padding=(10, 6))
        self.load_step_btn.grid(row=row, column=3, sticky="nsew", padx=(2, 0), pady=(0, 5))


    def start_ui_services(self):
        self.tooltip_manager = ImageTooltipManager(
            self.steps_tree, lambda: self.steps
        )

    @staticmethod
    def _validate_steps_shape(steps: Any) -> None:
        if not isinstance(steps, list):
            raise ValueError('steps must be a list')
        for index, step in enumerate(steps):
            if not isinstance(step, dict):
                raise ValueError(f'step {index + 1} must be a mapping')
            if 'action' not in step:
                raise ValueError(f"step {index + 1} is missing 'action'")
            if not isinstance(step.get('params'), dict):
                raise ValueError(f"step {index + 1} has invalid 'params'")

    def load_steps(self, steps: list[dict[str, Any]]) -> None:
        """Atomically replace state with a private copy of validated steps."""
        self._validate_steps_shape(steps)
        replacement = copy.deepcopy(steps)
        self.steps = replacement
        self.last_test_location = None
        self.last_test_locate_anchor = None
        self.editing_step_has_cache_box = False
        self.cancel_edit_mode()

    def clear_steps(self) -> None:
        self.steps = []
        self.last_test_location = None
        self.last_test_locate_anchor = None
        self.editing_step_has_cache_box = False
        self.cancel_edit_mode()

    def get_steps_snapshot(self) -> list[dict[str, Any]]:
        return copy.deepcopy(self.steps)

    def has_steps(self) -> bool:
        return bool(self.steps)

    def is_editing(self) -> bool:
        return self.editing_index is not None

    def count_enabled_action(self, action: str) -> int:
        return sum(
            1 for step in self.steps
            if step.get('action') == action and step.get('enabled', True)
        )

    def get_press_key_hotkey_conflicts(
            self,
            reserved_hotkeys: Mapping[str, str],
    ) -> list[tuple[int, str, str]]:
        """Return PRESS_KEY steps that would trigger a reserved app hotkey."""
        normalized_reserved = {
            HotkeyUtils.normalize_key_chord(configured): str(purpose)
            for purpose, configured in reserved_hotkeys.items()
            if HotkeyUtils.normalize_key_chord(configured)
        }
        conflicts = []
        for index, step in enumerate(self.steps, start=1):
            if not isinstance(step, dict) or step.get('action') != 'PRESS_KEY':
                continue
            params = step.get('params')
            if not isinstance(params, dict):
                continue
            key_chord = HotkeyUtils.normalize_key_chord(params.get('key', ''))
            purpose = normalized_reserved.get(key_chord)
            if purpose:
                conflicts.append((index, purpose, key_chord))
        return conflicts

    def set_available_ocr_keys(self, keys) -> None:
        """Update options used by forms built after this call."""
        self.available_ocr_keys = list(keys)

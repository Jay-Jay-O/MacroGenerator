import time
import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, messagebox
import threading
from autoclicker import AutoClicker

class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Macro Generator")
        self.root.geometry("340x600") 

        self.bot = AutoClicker()
        self.actions = []

        # --- Control Frame ---
        control_frame = ttk.LabelFrame(root, text="Settings", padding=10)
        control_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(control_frame, text="Loops:").grid(row=0, column=0, padx=5, sticky="e")
        self.loop_var = tk.IntVar(value=1)
        ttk.Entry(control_frame, textvariable=self.loop_var, width=10).grid(row=0, column=1, padx=5, sticky="w")

        ttk.Label(control_frame, text="Default Delay (ms):").grid(row=1, column=0, padx=5, sticky="e")
        self.delay_var = tk.IntVar(value=300)
        ttk.Entry(control_frame, textvariable=self.delay_var, width=10).grid(row=1, column=1, padx=5, sticky="w")

        # --- Tabs ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(padx=10, pady=5, fill="both", expand=True)

        # Tab 1: Actions List
        self.tab_actions = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_actions, text="Actions List")

        # Tab 2: Logs
        self.tab_logs = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_logs, text="Logs")

        # --- Action List UI (Tab 1) ---
        # Frame for Treeview and Scrollbar
        tree_frame = ttk.Frame(self.tab_actions)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Treeview
        columns = ("Type", "Details", "Duration")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("Type", text="Type")
        self.tree.heading("Details", text="Details")
        self.tree.heading("Duration", text="Duration (ms)")
        
        self.tree.column("Type", width=40, anchor="center")
        self.tree.column("Details", width=140, anchor="w")
        self.tree.column("Duration", width=80, anchor="center")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Buttons for List Management
        btn_action_frame = ttk.Frame(self.tab_actions)
        btn_action_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(btn_action_frame, text="Delete Selected", command=self.delete_action).pack(side="left", padx=2)
        ttk.Button(btn_action_frame, text="Clear All", command=self.reset_actions).pack(side="right", padx=2)
        ttk.Button(btn_action_frame, text="Edit Selected", command=self.edit_action).pack(side="left", padx=2)

        # --- Main Button Frame (Bottom) ---
        main_btn_frame = ttk.Frame(root, padding=10)
        main_btn_frame.pack(fill="x", padx=5)

        ttk.Button(main_btn_frame, text="Add Mouse Action", command=self.add_mouse_action).pack(fill="x", pady=2)
        ttk.Button(main_btn_frame, text="Add Key Action", command=self.add_key_action).pack(fill="x", pady=2)
        
        self.run_btn = ttk.Button(main_btn_frame, text="RUN MACRO", command=self.start_macro_thread)
        self.run_btn.pack(fill="x", pady=5)

        # --- Logs UI (Tab 2) ---
        self.log_area = scrolledtext.ScrolledText(self.tab_logs, width=70, height=20)
        self.log_area.pack(padx=10, pady=10, fill="both", expand=True)
        self.log("Ready to add actions...")
        
        # --- Drag & Drop Bindings ---
        self.tree.bind("<Button-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_release)
        self.tree.bind("<Double-1>", lambda e: self.edit_action())
        self.drag_item = None

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def refresh_list(self):
        # Clear current list
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Repopulate
        for i, action in enumerate(self.actions):
            t = action["type"]
            dur = action.get("duration", 0)
            details = ""
            
            if t == "mouse":
                btn = action["button"]
                if btn == 1: b_str = "Left"
                elif btn == 2: b_str = "Right"
                elif btn == 3: b_str = "Middle"
                elif btn == 4: b_str = "Scroll"
                else: b_str = "?"
                
                if btn == 4:
                    amount = action.get("scroll_amount", 0)
                    details = f"Scroll {amount}"
                elif action.get("drag", False):
                    details = f"Drag {b_str} ({action['x']},{action['y']}) -> ({action['end_x']},{action['end_y']})"
                else:
                    details = f"Click {b_str} at ({action['x']}, {action['y']})"
            
            elif t == "key":
                code = action["code"]
                name = self.bot.getKeyName(code)
                details = f"Press {name}"
                
            elif t == "shortcut":
                codes = action["codes"]
                names = [self.bot.getKeyName(k) for k in codes]
                details = f"Shortcut {' + '.join(names)}"
            
            self.tree.insert("", "end", iid=str(i), values=(t.upper(), details, f"{dur}ms"))

    def delete_action(self):
        sel = self.tree.selection()
        if not sel: return
        
        idx = int(sel[0])
        del self.actions[idx]
        self.refresh_list()
        self.log(f"Deleted action {idx+1}")

    # --- Drag & Drop Handlers ---
    def on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_item = item

    def on_drag_motion(self, event):
        # Optional: Change cursor or visual feedback
        pass

    def on_drag_release(self, event):
        if not self.drag_item: return
        
        target_item = self.tree.identify_row(event.y)
        if target_item and target_item != self.drag_item:
            # Reorder self.actions
            src_idx = int(self.drag_item)
            tgt_idx = int(target_item)
            
            # Move item in list
            item = self.actions.pop(src_idx)
            self.actions.insert(tgt_idx, item)
            
            self.refresh_list()
            self.log(f"Moved action from {src_idx+1} to {tgt_idx+1}")
            # Reselect the moved item
            self.tree.selection_set(str(tgt_idx))
            
        self.drag_item = None


    def add_mouse_action(self, edit_index=None):
        # Create custom popup
        popup = tk.Toplevel(self.root)
        popup.title("Edit Mouse Action" if edit_index is not None else "Add Mouse Action")
        popup.geometry("300x320")
        popup.transient(self.root)
        popup.grab_set()
        
        # Center relative to parent
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (300 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (320 // 2)
        popup.geometry(f"+{x}+{y}")
        
        # Default values
        def_btn = 1
        def_x = 0
        def_y = 0
        def_scroll = 120
        def_dur = 100
        def_drag = False
        def_ex = 0
        def_ey = 0
        
        # Load if editing
        if edit_index is not None:
             a = self.actions[edit_index]
             def_btn = a["button"]
             def_x = a["x"]
             def_y = a["y"]
             def_scroll = a.get("scroll_amount", 120)
             def_dur = a.get("duration", 100)
             def_drag = a.get("drag", False)
             def_ex = a.get("end_x", 0)
             def_ey = a.get("end_y", 0)

        ttk.Label(popup, text="Action Type:").pack(pady=2)
        btn_var = tk.IntVar(value=def_btn)
        
        # Function to toggle fields
        def toggle_fields():
            action = btn_var.get()
            if action == 4: # Scroll
                coord_frame.pack_forget()
                drag_check.pack_forget()
                scroll_frame.pack(pady=5, after=btn_frame)
            else: # Click
                scroll_frame.pack_forget()
                coord_frame.pack(pady=5, after=btn_frame)
                drag_check.pack(pady=5, after=coord_frame)
        
        btn_frame = ttk.Frame(popup)
        btn_frame.pack()
        ttk.Radiobutton(btn_frame, text="Left", variable=btn_var, value=1, command=toggle_fields).pack(side="left")
        ttk.Radiobutton(btn_frame, text="Right", variable=btn_var, value=2, command=toggle_fields).pack(side="left")
        ttk.Radiobutton(btn_frame, text="Mid", variable=btn_var, value=3, command=toggle_fields).pack(side="left")
        ttk.Radiobutton(btn_frame, text="Scroll", variable=btn_var, value=4, command=toggle_fields).pack(side="left")

        # --- Coordinates Frame (Container for Start/End) ---
        coord_frame = ttk.Frame(popup)
        coord_frame.pack(pady=5)

        ttk.Label(coord_frame, text="Move mouse and press:").pack(pady=2)
        ttk.Label(coord_frame, text="CTRL for Start | SHIFT for End (Drag)").pack(pady=0)
        
        # X/Y Frame (Start)
        xy_frame = ttk.Frame(coord_frame)
        xy_frame.pack(pady=2)
        
        ttk.Label(xy_frame, text="Start X:").grid(row=0, column=0)
        x_var = tk.IntVar(value=def_x)
        x_entry = ttk.Entry(xy_frame, textvariable=x_var, width=8)
        x_entry.grid(row=0, column=1, padx=5)

        ttk.Label(xy_frame, text="Start Y:").grid(row=0, column=2)
        y_var = tk.IntVar(value=def_y)
        y_entry = ttk.Entry(xy_frame, textvariable=y_var, width=8)
        y_entry.grid(row=0, column=3, padx=5)
        
        # X/Y Frame (End)
        end_frame = ttk.Frame(coord_frame)
        end_frame.pack(pady=2)
        
        ttk.Label(end_frame, text="End X:").grid(row=0, column=0)
        end_x_var = tk.IntVar(value=def_ex)
        end_x_entry = ttk.Entry(end_frame, textvariable=end_x_var, width=8)
        end_x_entry.grid(row=0, column=1, padx=5)

        ttk.Label(end_frame, text="End Y:").grid(row=0, column=2)
        end_y_var = tk.IntVar(value=def_ey)
        end_y_entry = ttk.Entry(end_frame, textvariable=end_y_var, width=8)
        end_y_entry.grid(row=0, column=3, padx=5)

        # Drag Checkbox
        drag_var = tk.BooleanVar(value=def_drag)
        drag_check = ttk.Checkbutton(popup, text="Drag / Hold to (Rectangle)", variable=drag_var)
        drag_check.pack(pady=5)
        
        # --- Scroll Frame ---
        scroll_frame = ttk.Frame(popup)
        # Don't pack initially (default is click)
        
        ttk.Label(scroll_frame, text="Scroll Amount (±120 = 1 notch):").pack(pady=2)
        scroll_var = tk.IntVar(value=def_scroll)
        ttk.Entry(scroll_frame, textvariable=scroll_var, width=10).pack()

        # Duration
        common_frame = ttk.Frame(popup)
        common_frame.pack(pady=5, side="bottom", fill="x") # Pack duration at bottom
        
        ttk.Label(common_frame, text="Hold Duration (ms):").pack(side="top")
        dur_var = tk.IntVar(value=def_dur)
        ttk.Entry(common_frame, textvariable=dur_var, width=10).pack(side="top")
        
        # Trigger toggle to set initial correct state
        toggle_fields()
        
        # Add Button
        btn_text = "Save Changes" if edit_index is not None else "Add Action"
        ttk.Button(common_frame, text=btn_text, command=lambda: on_add()).pack(pady=10, side="bottom")

        # Polling for Ctrl/Shift key
        def check_input():
            if not popup.winfo_exists(): return
            
            # 0x11 is VK_CONTROL (Start Point)
            if self.bot.isKeyPressed(0x11):
                mx, my = self.bot.getCursorPos()
                x_var.set(mx)
                y_var.set(my)
            
            # 0x10 is VK_SHIFT (End Point)
            if self.bot.isKeyPressed(0x10):
                mx, my = self.bot.getCursorPos()
                end_x_var.set(mx)
                end_y_var.set(my)
                if not drag_var.get():
                     drag_var.set(True)

            popup.after(50, check_input)
        
        check_input()

        def on_add():
            try:
                action = {
                    "type": "mouse",
                    "x": x_var.get(),
                    "y": y_var.get(),
                    "button": btn_var.get(),
                    "scroll_amount": scroll_var.get(),
                    "duration": dur_var.get(),
                    "drag": drag_var.get() if btn_var.get() != 4 else False, # Force false if scroll
                    "end_x": end_x_var.get(),
                    "end_y": end_y_var.get()
                }
                
                desc = f"Click ({x_var.get()}, {y_var.get()})"
                if action['drag']:
                    desc = f"Drag ({x_var.get()},{y_var.get()}) -> ({end_x_var.get()},{end_y_var.get()})"
                elif btn_var.get() == 4:
                     desc = f"Scroll {scroll_var.get()}"
                
                if edit_index is not None:
                    self.actions[edit_index] = action
                    self.log(f"Edited: Mouse {desc} Btn {btn_var.get()}")
                else:
                    self.actions.append(action)
                    self.log(f"Added: Mouse {desc} Btn {btn_var.get()}")
                
                self.refresh_list()
                popup.destroy()
            except ValueError:
                messagebox.showerror("Error", "Invalid coordinates")

    def add_key_action(self, edit_index=None):
        # Create custom popup
        popup = tk.Toplevel(self.root)
        popup.title("Edit Key/Shortcut" if edit_index is not None else "Add Key/Shortcut")
        popup.geometry("350x250")
        popup.transient(self.root)
        popup.grab_set()

        # Center relative to parent
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (350 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (250 // 2)
        popup.geometry(f"+{x}+{y}")
        
        # Default values
        def_dur = 100
        initial_keys = []
        
        if edit_index is not None:
            a = self.actions[edit_index]
            def_dur = a.get("duration", 100)
            if a["type"] == "shortcut":
                initial_keys = a["codes"]
            elif a["type"] == "key":
                initial_keys = [a["code"]]
        
        ttk.Label(popup, text="Click box and press keys (Sequence):").pack(pady=5)
        
        # Capture Entry using a StringVar to update text
        key_display_var = tk.StringVar(value="Click to set keys...")
        key_entry = tk.Entry(popup, textvariable=key_display_var, width=30, justify="center", font=("Consolas", 10))
        key_entry.pack(pady=5, ipady=5)
        
        # Instructions
        ttk.Label(popup, text="(Click away to stop capturing)").pack(pady=0)

        # Duration
        ttk.Label(popup, text="Hold Duration (ms):").pack(pady=5)
        dur_var = tk.IntVar(value=def_dur)
        dur_entry = ttk.Entry(popup, textvariable=dur_var, width=10)
        dur_entry.pack()

        # State
        captured_keys = list(initial_keys)
        
        def update_display():
            if not captured_keys:
                key_display_var.set("Click to set keys...")
            else:
                names = [self.bot.getKeyName(k) for k in captured_keys]
                key_display_var.set(" + ".join(names))

        # Initial display update if editing
        update_display()

        def check_key():
            if not popup.winfo_exists(): return
            
            # Check VK codes
            for vk in range(1, 255):
                if vk in [1, 2, 4]: continue # Skip mouse buttons
                # Skip generic modifiers (CTRL/LCTRL issues)
                if vk in [0x10, 0x11, 0x12]: continue
                
                if self.bot.isKeyPressed(vk):
                    if vk not in captured_keys:
                        captured_keys.append(vk)
                        update_display()
                        
            popup.after(50, check_key)
        
        check_key()
        
        # Clear Button
        def clear_keys():
            captured_keys.clear()
            update_display()
            key_entry.focus_set()
            
        ttk.Button(popup, text="Clear", command=clear_keys).pack(pady=2)

        def on_add():
            if not captured_keys:
                messagebox.showerror("Error", "No keys captured.")
                return
                
            action = {
                "type": "shortcut",
                "codes": list(captured_keys),
                "duration": dur_var.get()
            }
            desc = "Shortcut " + "+".join([self.bot.getKeyName(k) for k in captured_keys])
            
            if edit_index is not None:
                self.actions[edit_index] = action
                self.log(f"Edited: {desc} for {dur_var.get()}ms")
            else:
                self.actions.append(action)
                self.log(f"Added: {desc} for {dur_var.get()}ms")
                
            self.refresh_list()
            popup.destroy()

        btn_text = "Save Changes" if edit_index is not None else "Add Action"
        ttk.Button(popup, text=btn_text, command=on_add).pack(pady=10)

    def reset_actions(self):
        self.actions.clear()
        self.refresh_list() # Clear the list view too
        self.log_area.delete('1.0', tk.END)
        self.log("Actions cleared.")

    def edit_action(self):
        sel = self.tree.selection()
        if not sel: return
        
        idx = int(sel[0])
        action = self.actions[idx]
        
        if action["type"] == "mouse":
            self.add_mouse_action(edit_index=idx)
        elif action["type"] in ["key", "shortcut"]:
            self.add_key_action(edit_index=idx)

    def start_macro_thread(self):
        t = threading.Thread(target=self.run_macro)
        t.start()


    def run_macro(self):
        loops = self.loop_var.get()
        delay = self.delay_var.get()
        
        self.run_btn.config(state="disabled")
        self.log(f"--- Starting Macro ({loops} Loops) ---")
        self.log("Press PAUSE BREAK to Emergency Stop")
        
        time.sleep(1) 
        
        try:
            for l in range(loops):
                if self.bot.isKeyPressed(0x13): 
                    raise Exception("Emergency Stop Triggered!")

                self.log(f"Loop {l + 1}/{loops}")
                for i, action in enumerate(self.actions):
                    if self.bot.isKeyPressed(0x13): 
                        raise Exception("Emergency Stop Triggered!")

                    if action["type"] == "mouse":
                        # Check if it is a drag action
                        if action.get("drag", False):
                            dur = action.get("duration", 0)
                            self.bot.mouseDrag(
                                action["x"], action["y"],
                                action["end_x"], action["end_y"],
                                action["button"], dur
                            )
                            self.log(f"  Executed: Drag ({action['x']},{action['y']}) -> ({action['end_x']},{action['end_y']})")
                        elif action["button"] == 4:
                            # Scroll
                            amount = action.get("scroll_amount", 0)
                            self.bot.mouseScroll(amount)
                            self.log(f"  Executed: Mouse Scroll {amount}")
                        else:
                            self.bot.mouseMove(action["x"], action["y"])
                            # Use mapped duration or default 100
                            dur = action.get("duration", 100)
                            self.bot.clickMouse(action["button"], dur, delay)
                            self.log(f"  Executed: Mouse Click ({action['x']}, {action['y']}) for {dur}ms")
                    
                    elif action["type"] == "key":
                        dur = action.get("duration", 100)
                        self.bot.keyPress(action["code"], dur)
                        time.sleep(delay / 1000.0)
                        self.log(f"  Executed: Key Press {action['code']} for {dur}ms")
                        
                    elif action["type"] == "shortcut":
                        dur = action.get("duration", 100)
                        self.bot.shortcut(action["codes"], dur)
                        time.sleep(delay / 1000.0)
                        names = [self.bot.getKeyName(k) for k in action["codes"]]
                        self.log(f"  Executed: Shortcut {'+'.join(names)} for {dur}ms")
            
            self.log("--- Macro Finished ---")

        except Exception as e:
            self.log(f"Error: {e}")
        
        finally:
            self.run_btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroApp(root)
    root.mainloop()

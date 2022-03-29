import tkinter as tk
import json
from tkinter import filedialog
from PIL import Image, ImageTk
from image_manager import ImageManager


class DragDropListbox(tk.Listbox):
    """A Tkinter listbox with drag'n'drop reordering of entries."""

    def __init__(self, master, **kw):
        kw["selectmode"] = tk.SINGLE
        tk.Listbox.__init__(self, master, kw)
        self.bind("<Button-1>", self.setCurrent)
        self.bind("<B1-Motion>", self.shiftSelection)
        self.curIndex = None

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)

    def shiftSelection(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i + 1, x)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i - 1, x)
            self.curIndex = i


class GUI:
    def __init__(self):
        self._image_manager = ImageManager()
        self._max_width = 800
        self._max_height = 560
        self._init_window()

    def start(self):
        self._window.mainloop()

    def _init_window(self):
        self._window = tk.Tk()
        self._window.title("照片黏貼紀錄表")
        self._window.geometry("1200x600")
        self._create_menu()
        self._create_frames()
        self._create_set_report_title()
        self._create_add_delete_image_buttons()
        self._create_image_list_box()
        self._create_show_image_canvas()
        self._create_insert_date()
        self._set_page_up_down_button()

    def _create_menu(self):
        self._main_menu = tk.Menu(self._window)
        self._window.config(menu=self._main_menu)
        self._file_menu = tk.Menu(self._main_menu, tearoff=0)
        self._main_menu.add_cascade(label="File", menu=self._file_menu)
        self._file_menu.add_command(
            label="Save as Word report", command=self._save_report
        )
        self._file_menu.add_command(
            label="Import data as json", command=self._import_data
        )
        self._file_menu.add_command(
            label="Export data as json", command=self._export_data
        )

    def _create_frames(self):
        # arrange main frames
        self._left_frame = tk.Frame(self._window, width="400", height="600")
        self._right_frame = tk.Frame(self._window, width="800", height="600")
        self._left_frame.grid(column=0, row=0)
        self._right_frame.grid(column=1, row=0)

        # arrange righht frames
        self._show_image_frame = tk.Frame(self._right_frame, width="800", height="560")
        self._insert_date_frame = tk.Frame(self._right_frame, width="800", height="20")
        self._show_image_frame.grid(column=0, row=0)
        self._insert_date_frame.grid(column=0, row=1)

        # arrange left frames
        self._image_list_frame = tk.Frame(self._left_frame, width="400", height="560")
        self._set_report_title_frame = tk.Frame(
            self._left_frame, width="400", height="20"
        )
        self._add_delete_image_btn_frame = tk.Frame(
            self._left_frame, width="400", height="20"
        )
        self._image_list_frame.grid(column=0, row=0, stick="ew")
        self._set_report_title_frame.grid(column=0, row=1, sticky="ew")
        self._add_delete_image_btn_frame.grid(column=0, row=2, sticky="ew")

    def _create_add_delete_image_buttons(self):
        self._add_delete_image_btn_frame.grid_columnconfigure((0, 1), weight=1)
        self._add_image_btn = tk.Button(
            self._add_delete_image_btn_frame, text="Add Images", command=self._add_image
        )
        self._delete_image_btn = tk.Button(
            self._add_delete_image_btn_frame,
            text="Delete Images",
            command=self._delete_image,
        )
        self._add_image_btn.grid(column=0, row=0, sticky="ew")
        self._delete_image_btn.grid(column=1, row=0, sticky="ew")

    def _create_set_report_title(self):
        def update_report_title(*args):
            title = self._report_title_variable.get()
            self._image_manager.update_report_title(title)

        self._set_report_title_frame.grid_columnconfigure(0, weight=1)
        self._set_report_title_frame.grid_columnconfigure(1, weight=4)
        self._set_report_title_lable = tk.Label(
            self._set_report_title_frame, text="Report Title"
        )
        self._report_title_variable = tk.StringVar()
        self._set_report_title_entry = tk.Entry(
            self._set_report_title_frame,
            textvariable=self._report_title_variable,
        )
        self._report_title_variable.trace("w", update_report_title)
        self._set_report_title_lable.grid(column=0, row=0)
        self._set_report_title_entry.grid(column=1, row=0, sticky="ew")

    def _create_image_list_box(self):
        self._image_list_box = DragDropListbox(
            self._image_list_frame, height=32, width=50
        )
        self._image_list_box.grid(column=0, row=0, sticky="nsew")
        y_scrollbar = tk.Scrollbar(
            self._image_list_frame, command=self._image_list_box.yview
        )
        x_scrollbar = tk.Scrollbar(
            self._image_list_frame,
            orient="horizontal",
            command=self._image_list_box.xview,
        )
        self._image_list_box.config(
            yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set
        )
        y_scrollbar.grid(row=0, column=1, rowspan=2, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")

        self._image_list_box.bind("<<ListboxSelect>>", self._select_image)
        self._image_list_box.bind("<Down>", self._click_down)
        self._image_list_box.bind("<Up>", self._click_up)
        self._image_list_box.bind("<Delete>", self._delete_image)

    def _create_show_image_canvas(self):
        self._image_canvas = tk.Canvas(
            self._show_image_frame, height="560", width="800", bg="white"
        )
        self._image_canvas.grid(row=0, column=0, sticky="nsew")

    def _create_insert_date(self):
        self._year_label = tk.Label(self._insert_date_frame, text="年").grid(
            column=1, row=0
        )
        self._month_label = tk.Label(self._insert_date_frame, text="月").grid(
            column=3, row=0
        )
        self._day_label = tk.Label(self._insert_date_frame, text="日").grid(
            column=5, row=0
        )
        self._hour_label = tk.Label(self._insert_date_frame, text="時").grid(
            column=7, row=0
        )
        self._minute_label = tk.Label(self._insert_date_frame, text="分").grid(
            column=9, row=0
        )
        self._second_label = tk.Label(self._insert_date_frame, text="秒").grid(
            column=11, row=0
        )

        def update_time(time_unit, *args):
            enrty_mapping = {
                "year": self._year_entry,
                "month": self._month_entry,
                "day": self._day_entry,
                "hour": self._hour_entry,
                "minute": self._minute_entry,
                "second": self._second_entry,
            }
            cur_index = self._image_list_box.curIndex
            file_name = self._image_list_box.get(cur_index)
            if cur_index is not None:
                value = enrty_mapping[time_unit].get()
                self._image_manager.update_time(file_name, time_unit, value)

        self._year_varialbe = tk.StringVar()
        self._month_varialbe = tk.StringVar()
        self._day_varialbe = tk.StringVar()
        self._hour_varialbe = tk.StringVar()
        self._minute_varialbe = tk.StringVar()
        self._second_varialbe = tk.StringVar()

        self._year_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._year_varialbe
        )
        self._year_entry.bind("<FocusOut>", lambda *args: update_time("year", *args))
        self._year_entry.grid(column=0, row=0)

        self._month_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._month_varialbe
        )
        self._month_entry.bind("<FocusOut>", lambda *args: update_time("month", *args))
        self._month_entry.grid(column=2, row=0)

        self._day_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._day_varialbe
        )
        self._day_entry.bind("<FocusOut>", lambda *args: update_time("day", *args))
        self._day_entry.grid(column=4, row=0)

        self._hour_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._hour_varialbe
        )
        self._hour_entry.bind("<FocusOut>", lambda *args: update_time("hour", *args))
        self._hour_entry.grid(column=6, row=0)

        self._minute_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._minute_varialbe
        )
        self._minute_entry.bind(
            "<FocusOut>", lambda *args: update_time("minute", *args)
        )
        self._minute_entry.grid(column=8, row=0)

        self._second_entry = tk.Entry(
            self._insert_date_frame, width=3, textvariable=self._second_varialbe
        )
        self._second_entry.bind(
            "<FocusOut>", lambda *args: update_time("second", *args)
        )
        self._second_entry.grid(column=10, row=0)

        def update_use_image_time(*ars):
            cur_index = self._image_list_box.curIndex
            file_name = self._image_list_box.get(cur_index)
            if cur_index is not None:
                value = self._is_use_image_time.get()
                self._image_manager.update_use_image_time(file_name, value)

        self._is_use_image_time = tk.BooleanVar()
        self._is_use_image_time.set(False)
        self._use_image_time_checkbox = tk.Checkbutton(
            self._insert_date_frame, text="Use Image Time", var=self._is_use_image_time
        ).grid(column=13, row=0)
        self._is_use_image_time.trace("w", update_use_image_time)

    def _add_image(self, event=None):
        files = filedialog.askopenfilenames(parent=self._window, title="Choose image")
        file_path = self._window.tk.splitlist(files)
        self._image_manager.add_image(file_path)
        for path in file_path:
            self._image_list_box.insert(tk.END, path)

    def _delete_image(self, event=None):
        cur_index = self._image_list_box.curIndex
        file_path = self._image_list_box.get(cur_index)
        self._image_manager.delete_iamge(file_path)
        self._image_list_box.delete(cur_index)
        self._image_canvas.delete("all")

    def _select_image(self, envet):
        self._show_image()
        self._show_time_info()

    def _show_time_info(self):
        cur_index = self._image_list_box.curIndex
        file_path = self._image_list_box.get(cur_index)
        cur_image_info = self._image_manager.get_image_info(file_path)
        self._year_varialbe.set(cur_image_info["time"]["year"])
        self._month_varialbe.set(cur_image_info["time"]["month"])
        self._day_varialbe.set(cur_image_info["time"]["day"])
        self._hour_varialbe.set(cur_image_info["time"]["hour"])
        self._minute_varialbe.set(cur_image_info["time"]["minute"])
        self._second_varialbe.set(cur_image_info["time"]["second"])
        self._is_use_image_time.set(cur_image_info["use_image_time"])

    def _show_image(self):
        def resize_to_fit_height(img_width, img_height):
            img_width = int(img_width * (self._max_height / img_height))
            img_height = self._max_height
            return img_width, img_height

        def resize_to_fit_width(img_width, img_height):
            img_height = int(img_height * (self._max_width / img_width))
            img_width = self._max_width
            return img_width, img_height

        index = self._image_list_box.curIndex
        filename = self._image_list_box.get(index)
        img = Image.open(filename)
        img_width, img_height = img.size
        height_ratio = None
        if img_height > self._max_height:
            height_ratio = img_height / self._max_height
        width_ratio = None
        if img_width > self._max_width:
            width_ratio = img_width / self._max_width

        if height_ratio and width_ratio:
            if height_ratio > width_ratio:
                img_width, img_height = resize_to_fit_height(img_width, img_height)
            else:
                img_width, img_height = resize_to_fit_width(img_width, img_height)
        elif height_ratio:
            img_width, img_height = resize_to_fit_height(img_width, img_height)
        elif width_ratio:
            img_width, img_height = resize_to_fit_width
        img = img.resize((img_width, img_height), Image.ANTIALIAS)
        img = ImageTk.PhotoImage(img)
        self._image_canvas.image = img
        self._image_canvas.create_image(
            self._max_width / 2, self._max_height / 2, image=img, anchor=tk.CENTER
        )

    def _save_report(self):
        save_file_name = filedialog.asksaveasfile(mode="wb", defaultextension=".docx")
        file_path_list = self._image_list_box.get(0, tk.END)
        self._image_manager.generate_report(save_file_name, file_path_list)

    def _export_data(self):
        save_file_name = filedialog.asksaveasfilename(defaultextension=".json")
        file_path_list = self._image_list_box.get(0, tk.END)
        sorted_image_date = self._image_manager.export_data(file_path_list)

        with open(save_file_name, "w") as fp:
            json.dump(sorted_image_date, fp, indent=4)

    def _import_data(self):
        import_data_path = filedialog.askopenfilename(
            parent=self._window, title="Choose Import Data"
        )
        print(f"Import data path: {import_data_path}")
        with open(import_data_path) as fp:
            try:
                import_data = json.load(fp)
            except:
                print("Please insert json file")

        self._image_list_box.delete(0, tk.END)
        try:
            for data in import_data:
                self._image_list_box.insert(tk.END, data["file_path"])
        except KeyError:
            print("Import data format error")

        self._image_manager.import_data(import_data)

    def _set_page_up_down_button(self):
        self._window.bind("<Prior>", self._click_up)
        self._window.bind("<Next>", self._click_down)

    def _click_down(self, event):
        list_box_size = self._image_list_box.size()
        cur_index = self._image_list_box.curIndex
        if cur_index is not None and cur_index < list_box_size - 1:
            self._image_list_box.curIndex = cur_index + 1
            self._select_image(None)

    def _click_up(self, event):
        cur_index = self._image_list_box.curIndex
        if cur_index is not None and cur_index > 0:
            self._image_list_box.curIndex = cur_index - 1
            self._select_image(None)


if __name__ == "__main__":
    print("Report Generator is starting....")
    GUI = GUI()
    GUI.start()

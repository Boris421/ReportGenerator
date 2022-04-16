from copy import deepcopy
from report_controller import ReportController


class ImageManager:
    data_template = {
        "time": {
            "year": "",
            "month": "",
            "day": "",
            "hour": "",
            "minute": "",
            "second": "",
        },
        "use_image_time": False,
        "rotate_image": False,
    }

    def __init__(self):
        self._report_title = ""
        self._image_data = {}

    def _sort_image_data(self, file_path_list):
        sorted_image_date = []
        for file_path in file_path_list:
            image_data = self._image_data[file_path]
            image_data["file_path"] = file_path
            sorted_image_date.append(image_data)

        return sorted_image_date

    def add_image(self, file_paths):
        for path in file_paths:
            data = deepcopy(self.data_template)
            self._image_data[path] = data
        print(f"Add images {self._image_data}")

    def delete_iamge(self, file_path):
        del self._image_data[file_path]
        print(f"Delete image {self._image_data}")

    def get_image_info(self, file_path):
        return self._image_data[file_path]

    def generate_report(self, file_name, file_path_list):
        report_generator = ReportController()
        sorted_image_date = self._sort_image_data(file_path_list)

        print("===== Start tp Generate Report =====")
        report_generator.set_data(sorted_image_date, self._report_title)
        print("===== Generating Report =====")
        report_generator.generate_doc()
        print("===== Saving Report =====")
        report_generator.save(file_name)
        print("===== Generate Report Finished =====")
        report_generator.clear()

    def export_data(self, file_path_list):
        sorted_image_date = self._sort_image_data(file_path_list)
        return sorted_image_date

    def import_data(self, import_data):
        organized_image_data = {}
        try:
            for data in import_data:
                tmp_data = {}
                for key in data:
                    if key in ["time", "use_image_time", "rotate_image"]:
                        tmp_data[key] = data[key]
                organized_image_data[data["file_path"]] = tmp_data
        except:
            print("Import data format error")
        self._image_data = organized_image_data

    def update_time(self, file_path, time_unit, value):
        self._image_data[file_path]["time"][time_unit] = value

    def update_use_image_time(self, file_path, value):
        self._image_data[file_path]["use_image_time"] = value

    def update_report_title(self, title):
        self._report_title = title

    def update_rotate_image(self, file_path, value):
        self._image_data[file_path]["rotate_image"] = value

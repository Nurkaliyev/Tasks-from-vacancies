import os
import pandas

class Filepath_test_assignment:
    def _get_files(self):
        for root, dirs, files in os.walk(os.getcwd()):
            for file in files:
                if file.startswith('.'):
                    file = ' ' + file
                yield os.path.join(root, file)

    def _get_file_info(self, files):
        data = []
        for file in files:
            folder, filename = os.path.split(file)
            name, ext = os.path.splitext(filename)
            data.append((len(data) + 1, folder, name, ext))
        return data

    def _save_to_excel(self, data):
        df = pandas.DataFrame(data, columns=['Number', 'Folder', 'Name', 'Extension'])
        df.to_excel('result.xlsx', engine='openpyxl', index=False)

    def run(self):
        files = self._get_files()
        data = self._get_file_info(files)
        self._save_to_excel(data)

if __name__ == "__main__":
    task = Filepath_test_assignment()
    task.run()
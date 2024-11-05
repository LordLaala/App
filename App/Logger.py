import os

class Logger:

    def __init__(self,path="C:\\Users\\BAA3012\\PycharmProjects\\pythonProject1",path_name='log.txt'):
        self.path_name = path_name
        self.path = path+f"\\{self.path_name}"
        if os.path.exists(self.path) and os.path.isfile(self.path):
            os.remove(self.path)
            open(self.path, "x").close()
        else:
            open(self.path, "x").close()

    def write_error_messages(self,filepath,name,reason,codepart):
        with open(self.path,"a") as file:
            file.write(f'Filepath:\n{filepath}\n{name}\nReason:{reason}\nCodepart{codepart}')
            file.close()

    def write_error_json_ZI(self,name):
        with open(self.path,"a") as file:
            file.write(f'Nur Zusaetzliche Informationen\n{name}\n')
            file.close()
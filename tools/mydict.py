class MyDict:
    def __init__(self, data):
        self.data = data

    # 通过名称获取对应内容
    def get(self, name):
        if name in self.data:
            return self.data[name][0]
        else:
            return None

    # 添加内容到字典
    def add(self, name, content):
        if name in self.data:
            self.data[name].append(content)
        else:
            self.data[name] = [content]

# # 示例用法
# my_dict = MyDict({"名称1": ["内容1", "内容2", "内容3"], "名称2": ["内容4", "内容5", "内容6"]})
# my_dict.add("名称1", "新内容")
# print(my_dict.get("名称1"))  # 输出 ["内容1", "内容2", "内容3", "新内容"]

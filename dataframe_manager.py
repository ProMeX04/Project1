import pandas
from io import StringIO


class DataFrameManager:
    def __init__(self, response, current_lesson, lesson, keyword):
        self.df = pandas.read_html(StringIO(response.text))[0]

        self.keyword = keyword
        self.lesson = lesson
        self.total_lesson = len(self.lesson)
        self.current_lesson = current_lesson

        self.df = self.df.drop(columns=["Organization", "Points"])

        self.df.iloc[:, 2:] = self.df.iloc[:, 2:].apply(
            lambda x: x.astype(str).str[0:3])
        self.list_point = list(
            map(lambda x: x.split()[1], self.df.columns.tolist()[2:]))

    def cal_point_for_each_person(self, order):
        begin = 2
        person_point = []
        for index, lesson_i in enumerate(self.lesson):
            count_ac = 0
            count_wa = 0
            for j in range(begin, begin + lesson_i):
                if str(self.df.iloc[order, j]).isdigit():
                    if self.df.iloc[order, j] == self.list_point[j - 2]:
                        count_ac += 1
                    else:
                        count_wa += 1
            if count_ac or count_wa or index < self.current_lesson and lesson_i != 0:
                person_point.append(f"A:{count_ac}  W:{count_wa}")
            else:
                person_point.append("")
            begin += lesson_i

        return person_point

    def new_df(self):
        rows = []
        for i in range(self.df.shape[0]):
            name = self.df.iloc[i, 1].split()[0]
            row = [i+1, name]
            if self.keyword not in name.lower():
                continue
            rows.append([i+1, name] + self.cal_point_for_each_person(i))
        return pandas.DataFrame(rows, columns=["Rank", "Name"] + [f"Lesson {i}: {self.lesson[i-1]}" for i in range(1, self.total_lesson + 1)])

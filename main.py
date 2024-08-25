import openpyxl
import dataclasses
import typing as t
import pathlib
from selenium import webdriver
from glob import glob
from time import sleep

# CONST
WB = openpyxl.load_workbook("./Rozklad_23_24.xlsx")
SHEET_FOR_PARSE = WB["Розклад"]


@dataclasses.dataclass
class InfoAboutLesson:
    numerator: str | None
    numerator_audience: str | None
    denominator: str | None
    denominator_audience: str | None

    @classmethod
    def from_dict(cls, list_: list) -> t.Self:
        return cls(
            numerator=list_[0][0],
            numerator_audience=list_[0][1],
            denominator=list_[1][0],
            denominator_audience=list_[1][1],
        )


@dataclasses.dataclass
class InfoAboutDay:
    day: str
    first: InfoAboutLesson | None
    second: InfoAboutLesson | None
    third: InfoAboutLesson | None
    fourth: InfoAboutLesson | None
    fifth: InfoAboutLesson | None
    sixth: InfoAboutLesson | None

    @classmethod
    def from_dict(cls, list_: list[str], day: str) -> t.Self:
        return cls(
            day=day,
            first=InfoAboutLesson.from_dict(list_[0:2]),
            second=InfoAboutLesson.from_dict(list_[2:4]),
            third=InfoAboutLesson.from_dict(list_[4:6]),
            fourth=InfoAboutLesson.from_dict(list_[6:8]),
            fifth=InfoAboutLesson.from_dict(list_[8:10]),
            sixth=InfoAboutLesson.from_dict(list_[10:12]),
        )


@dataclasses.dataclass
class Schedule:
    group_name: str
    monday: InfoAboutDay | None
    tuesday: InfoAboutDay | None
    wednesday: InfoAboutDay | None
    thursday: InfoAboutDay | None
    friday: InfoAboutDay | None

    @classmethod
    def from_dict(cls, list_: list) -> t.Self:
        return cls(
            group_name=list_[0][0],
            monday=InfoAboutDay.from_dict(list_[1:13], "Понеділок"),
            tuesday=InfoAboutDay.from_dict(list_[13:25], "Вівторок"),
            wednesday=InfoAboutDay.from_dict(list_[25:37], "Середа"),
            thursday=InfoAboutDay.from_dict(list_[37:49], "Четвер"),
            friday=InfoAboutDay.from_dict(list_[49:61], "П'ятниця"),
        )


def get_data_from_sheet(column: int) -> list[list[str]]:
    result = []
    column = column * 2 + 1
    start_row = 6
    end_row = 67

    for row in range(start_row, end_row):
        value1 = SHEET_FOR_PARSE.cell(row=row, column=column).value
        value2 = SHEET_FOR_PARSE.cell(row=row, column=column + 1).value
        result.append([value1, value2])

    return result


def transform_list_to_html_list(list_: list[list[str]]) -> list[list[str]]:
    result = []
    for count, row in enumerate(list_):
        # 2 gray ... 2 white etc.
        color = "gray" if not ((count - 1) // 2) % 2 else "white"
        middle_clear = f"<td class='middle clear {color}'></td>"
        end_clear = f"<td class='end clear {color}'></td>"
        middle = f"<td class='middle {color}'>{row[0]}</td>"
        end = f"<td class='end {color}'>{row[1]}</td>"
        if count == 0:
            result.append([row[0], row[1]])
        elif row[0] is not None and row[1] is not None:
            result.append([middle, end])
        elif row[0] is not None and row[1] is None:
            result.append([middle, end_clear])
        elif row[0] is None and row[1] is None:
            result.append([middle_clear, end_clear])
        elif row[0] is None and row[1] is not None:
            result.append([middle_clear, end])
    return result


def cut_big_words(list_: list[list[str]]) -> list[list[str]]:
    change_on_it = {
        "Мистецтво (осн реставр)": "Мистецтво (осн реста)",
        "Англійська мова (проф)": "Англ мова (проф)",
        "Дизайн (Поліграфія) ІІ": "Дизайн (Поліграф.) ІІ",
        "Інформ техн і комп граф": "ІТ і Комп граф",
        "Встеп до спец (Технології)": "Вступ до спец (Техн)",
        "Метод фіз / Метод технол": "Метод фіз / технол",
        "Метод технол / Метод фіз вих": "Метод техн / фіз виховання",
        "Туристична і краєзнавча": "Туристична і краєзнав",
        "Польська / Інформаційні технолог": "Польска / Інформ техн",
        "Інформаційні / Польська": "Інформ / Польська",
        "Захист України (дів,юн)": "Зах. України (дів,юн)",
    }
    result = []
    for row in list_:
        cut_words = []
        for word in row:
            if len(f"{word}") > 21:
                try:
                    cut_words.append(change_on_it[f"{word}"])
                except:
                    cut_words.append(word)
            else:
                cut_words.append(word)
        result.append(cut_words)
    return result


def get_finished_schedule_object(column: int) -> Schedule:
    raw_data = get_data_from_sheet(column)
    fried_data = cut_big_words(raw_data)
    html_format = transform_list_to_html_list(fried_data)
    return Schedule.from_dict(html_format)


def get_first_block(element: str, lesson: int, schedules: Schedule):
    return f"""
                <tr>
                    <td class="number-lesson start" rowspan="2">
                        {lesson}
                    </td>
                    {getattr(schedules.monday, element).numerator}
                    {getattr(schedules.monday, element).numerator_audience}
                    {getattr(schedules.tuesday, element).numerator}
                    {getattr(schedules.tuesday, element).numerator_audience}
                    {getattr(schedules.wednesday, element).numerator}
                    {getattr(schedules.wednesday, element).numerator_audience}
                </tr>
                <tr>                    
                    {getattr(schedules.monday, element).denominator}
                    {getattr(schedules.monday, element).denominator_audience}
                    {getattr(schedules.tuesday, element).denominator}
                    {getattr(schedules.tuesday, element).denominator_audience}
                    {getattr(schedules.wednesday, element).denominator}
                    {getattr(schedules.wednesday, element).denominator_audience}
                </tr>
    """


def get_second_block(element: str, lesson: int, schedules: Schedule):
    return f"""
                <tr>
                    <td class="number-lesson start" rowspan="2">
                        {lesson}
                    </td>
                    {getattr(schedules.thursday, element).numerator}
                    {getattr(schedules.thursday, element).numerator_audience}
                    {getattr(schedules.friday, element).numerator}
                    {getattr(schedules.friday, element).numerator_audience}
                    {getattr(schedules.monday, element).numerator}
                    {getattr(schedules.monday, element).numerator_audience}
                </tr>
                <tr>                    
                    {getattr(schedules.thursday, element).denominator}
                    {getattr(schedules.thursday, element).denominator_audience}
                    {getattr(schedules.friday, element).denominator}
                    {getattr(schedules.friday, element).denominator_audience}
                    {getattr(schedules.monday, element).denominator}
                    {getattr(schedules.monday, element).denominator_audience}
                </tr>
    """


def import_data_to_html(schedules: Schedule):
    text = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Title</title>
        <link rel="stylesheet" href="../style.css">
    </head>
    <body>
        <div class="center">
            <table>
                <tr class="week">
                    <td class="start">1А</td>
                    <td class="week-name" colspan="2">{schedules.monday.day}</td>
                    <td class="week-name" colspan="2">{schedules.tuesday.day}</td>
                    <td class="week-name" colspan="2">{schedules.wednesday.day}</td>
                </tr>

                {get_first_block("first", 1, schedules)}
                {get_first_block("second", 2, schedules)}
                {get_first_block("third", 3, schedules)}
                {get_first_block("fourth", 4, schedules)}
                {get_first_block("fifth", 5, schedules)}
                {get_first_block("sixth", 6, schedules)}


                <tr class="week">
                    <td class="start">1А</td>
                    <td class="week-name" colspan="2">{schedules.thursday.day}</td>
                    <td class="week-name" colspan="2">{schedules.friday.day}</td>
                    <td class="week-name" colspan="2">{schedules.monday.day}</td>
                </tr>
                {get_second_block("first", 1, schedules)}
                {get_second_block("second", 2, schedules)}
                {get_second_block("third", 3, schedules)}
                {get_second_block("fourth", 4, schedules)}
                {get_second_block("fifth", 5, schedules)}
                {get_second_block("sixth", 6, schedules)}
            </table>
        </div>

    </body>
    </html>
        """

    pathlib.Path(f"./variant/").mkdir(parents=True, exist_ok=True)
    with open(f"./variant/{schedules.group_name}.html", "w") as file:
        file.write(text)


def parse_all_schedules(count: int):
    # к-ть розкладів
    for i in range(1, count + 1):
        schedules = get_finished_schedule_object(i)
        import_data_to_html(schedules)


def parse_all_schedules_to_photo():
    driver = webdriver.Firefox()
    for i in range(37):
        group_name = glob("./variant/*", recursive=True)[i]
        html_file = "file://" + f"/home/salo/huta/{group_name}"
        driver.get(html_file)
        driver.set_window_size(height=1200, width=1000)
        pathlib.Path(f"./screenshots").mkdir(parents=True, exist_ok=True)
        driver.save_screenshot(f"./screenshots/{group_name[9:]}.png")
        sleep(10)
    # driver.quit()


if "__main__" == __name__:
    parse_all_schedules(38)
    parse_all_schedules_to_photo()

import openpyxl

# from openpyxl.cell.cell import Cell
import dataclasses
import typing as t
from pprint import pprint
from random import choice
import pathlib
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import time
from glob import glob


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
    first: InfoAboutLesson | None
    second: InfoAboutLesson | None
    third: InfoAboutLesson | None
    fourth: InfoAboutLesson | None
    fifth: InfoAboutLesson | None
    sixth: InfoAboutLesson | None

    @classmethod
    def from_dict(cls, list_: list[str]) -> t.Self:
        return cls(
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
            monday=InfoAboutDay.from_dict(list_[1:13]),
            tuesday=InfoAboutDay.from_dict(list_[13:25]),
            wednesday=InfoAboutDay.from_dict(list_[25:37]),
            thursday=InfoAboutDay.from_dict(list_[37:49]),
            friday=InfoAboutDay.from_dict(list_[49:61]),
        )


def get_data_from_sheet(wb, column: int):
    sheet_for_parse = wb["Розклад"]
    row = 6
    result = []
    column = 2 * column + 1

    for row in range(row, 67):
        value1 = sheet_for_parse.cell(row=row, column=column).value
        value2 = sheet_for_parse.cell(row=row, column=column + 1).value

        finish1 = f"""<td class="middle">
                    {value1}
                </td>"""
        finish2 = f"""<td class="end">
                    {value2}
                </td>"""
        if row == 6:
            print(value1)
            result.append([value1, value2])
        elif value1 is not None and value2 is not None:
            # print(row, value1, value2)
            result.append([finish1, finish2])
        elif value1 is not None and value2 is None:
            # print(row, value1, value2)
            result.append([finish1, "<td class='end'></td>"])
        elif value1 is None and value2 is None:
            result.append(
                ["<td class='middle clear'></td>", "<td class='end clear'></td>"]
            )
        elif value1 is None and value2 is not None:
            result.append(["<td class='middle clear'></td>", finish2])
    return Schedule.from_dict(result)


def import_data_to_html(schedules0: Schedule, day: str):
    schedules = schedules0.monday
    match day:
        case "Понеділок":
            schedules: InfoAboutDay = schedules0.monday
        case "Вівторок":
            schedules: InfoAboutDay = schedules0.tuesday
        case "Середа":
            schedules: InfoAboutDay = schedules0.wednesday
        case "Четвер":
            schedules: InfoAboutDay = schedules0.thursday
        case "П'ятниця":
            schedules: InfoAboutDay = schedules0.friday

    text = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Title</title>
    <link rel="stylesheet" href="../style.css">
</head>
<body>
    <!-- tr th td   -->
    <div class="center">
        <table>
            <tr class="week">
                <td class="week-name" colspan="3">{day}</td>
            </tr>
            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    1
                </td>
                {schedules.first.numerator}
                {schedules.first.numerator_audience}
            </tr>
            <tr class="first-lesson">                    
                {schedules.first.denominator}
                {schedules.first.denominator_audience}
            </tr>


            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    2
                </td>
                {schedules.second.numerator}
                {schedules.second.numerator_audience}
            </tr>
            <tr class="first-lesson">
                {schedules.second.denominator}
                {schedules.second.denominator_audience}
            </tr>


            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    3
                </td>
                {schedules.third.numerator}
                {schedules.third.numerator_audience}
            </tr>
            <tr class="first-lesson">
                {schedules.third.denominator}
                {schedules.third.denominator_audience}
            </tr>


            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    4
                </td>
                {schedules.fourth.numerator}
                {schedules.fourth.numerator_audience}
            </tr>
            <tr class="first-lesson">
                {schedules.fourth.denominator}
                {schedules.fourth.denominator_audience}
            </tr>


            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    5
                </td>
                {schedules.fifth.numerator}
                {schedules.fifth.numerator_audience}
            </tr>
            <tr class="first-lesson">
                {schedules.fifth.denominator}
                {schedules.fifth.denominator_audience}
            </tr>


            <tr class="first-lesson">
                <td class="number-lesson start" rowspan="2">
                    6
                </td>
                {schedules.sixth.numerator}
                {schedules.sixth.numerator_audience}
            </tr>
            <tr class="first-lesson">
                {schedules.sixth.denominator}
                {schedules.sixth.denominator_audience}
            </tr>
        </table>
    </div>

</body>
</html>
    """
    pathlib.Path(f"./variant/{schedules0.group_name}").mkdir(
        parents=True, exist_ok=True
    )
    with open(f"./variant/{schedules0.group_name}/{day}.html", "w") as file:
        file.write(text)


def parse_all_schedules(count: int):
    wb = openpyxl.load_workbook("./Rozklad_23_24.xlsx")
    # к-ть розкладів
    for i in range(1, count + 1):
        print(i)
        for day in ["Понеділок", "Вівторок", "Середа", "Четвер", "П'ятниця"]:
            print(day)
            schedules = get_data_from_sheet(wb, i)
            import_data_to_html(schedules, day=day)


def parse_all_schedules_to_photo():
    driver = webdriver.Firefox()
    for i in range(38):
        group_name = glob("./variant/*/", recursive=True)[i]
        for day in range(5):
            path_to_photo_by_group = glob(f"{group_name}*", recursive=True)[day]
            html_file = "file://" + f"/home/salo/huta/{path_to_photo_by_group[1:]}"
            driver.get(html_file)
            driver.set_window_size(height=776, width=714)
            pathlib.Path(f"./screenshots/{group_name[10:]}").mkdir(
                parents=True, exist_ok=True
            )
            driver.save_screenshot(
                f"./screenshots/{group_name[10:]}/{path_to_photo_by_group[13:-5]}.png"
            )
    driver.quit()


if "__main__" == __name__:
    parse_all_schedules(38)
    parse_all_schedules_to_photo()

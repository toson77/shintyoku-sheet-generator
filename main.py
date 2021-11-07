from calendar import monthcalendar
import gen_sheet as gs


def main():
    B4: tuple[str, ...] = ("太郎1", "太郎2")
    B4_2: tuple[str, ...] = ("太郎1", "太郎2", "太郎3")
    M1: tuple[str, ...] = ("太郎1", "太郎2")
    M2: tuple[str, ...] = ("太郎1", "太郎2")
    legends: dict[str, tuple] = {"B4": B4, "M1": M1, "M2": M2}
    legends2: dict[str, tuple] = {"B4": B4_2, "M1": M1, "M2": M2}
    ex_sheet1 = gs.GenSheet(legends=legends, year=2021,
                            month=11, filename="test")
    ex_sheet2 = gs.GenSheet(legends=legends2, year=2021,
                            month=12, filename="test2")
    ex_sheet1.gen()
    ex_sheet2.gen()


if __name__ == '__main__':
    main()

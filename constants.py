class CONSTANTS:
    APP_TITLE = 'Пересчёт вентилятора системы дымоудаления'
    MENU = ('Файл', 'Руководство', 'О программе')
    FILE_SUBMENU = ('Открыть', 'Сохранить', 'Сохранить как...', 'Экспорт в DOCX')
    ABOUT = '''<html>
    📄 Программа основана на методике, изложенной в приложении Б ГОСТ 53300-2009<hr>
    💡 Идея: Константин @ <font color="blue">nedich@mail.ru</font><br>
    🛠 Реализация: Юрий @ <font color="blue">akudja.technology@gmail.com</font><hr>
    Почта для предложений и замечаний: @ <font color="blue">akudja.technology@gmail.com</font><hr>
    🤑 Можете угостить нас 🍺: ХХХХ ХХХХ ХХХХ ХХХХ
    </html>
    '''
    NUM_ROWS = 15
    NUM_COLS = 5
    TABLE_COLUMN_WIDTH = {
        0:910,
        1:450,
        2:60,
        3:80,
        4:60,
    }
    TABLE_ROW_HEIGHT = 40


    class BUTTONS:
        ADD = 'Добавить\nучасток'
        DELETE = 'Удалить\nучасток'
        TOOLTIPS = (
            'Копировать последнюю\nтаблицу',
            'Добавить таблицу',
            'Удалить таблицу',
        )
        ICONS = (
            'copy.png',
            'add.png',
            'delete.png',
        )
        STYLES = (
            'QPushButton { border-radius: 20px; background-color: #EFEFEF; } QPushButton:hover { border: 2px solid grey; background-color: #99CCFF }',
            'QPushButton { border-radius: 20px; background-color: #EFEFEF; } QPushButton:hover { border: 2px solid grey; background-color: #99FF99 }',
            'QPushButton { border-radius: 20px; background-color: #EFEFEF; } QPushButton:hover { border: 2px solid grey; background-color: #FF9999 }',
        )


    class TABLE1:
        NAME = 'table_1'
        ROWS = 10
        COLUMNS = 5
        EDITABLE_ROWS = (0, 1, 3, 5, 7)
        SPAN_ROWS = (1, 3, 5)
        HEADER_ROWS = (0, 1, 3, 5, 7, 8, 9)
        HEADER = (
            ' Приведенное статическое давление вентилятора',
            ' Установленная при проектировании испытуемой системы вытяжной противодымной вентиляции значение температуры продуктов горения,\n непосредственно удаляемых из защищаемого помещения (коридора)',
            '',
            ' Температура воздуха в помещениях и в вытяжном вентиляционном канале при проведении приемо-сдаточных и периодических испытаний',
            '',
            ' Установленная при проектировании испытуемой системы вытяжной противодымной вентиляции значение температуры продуктов горения,\n перемещаемых вентилятором',
            '',
            ' Разность уровней фактического расположения входного устройства вентилятора и открытого дымоприемного устройства вытяжного\n канала',
            ' Плотность воздуха при температуре Ta',
            ' Плотность удаляемого газа, перемещаемого вентилятором (при температуре Tv)',
        )
        FORMULAS = (
            '-',
            '-',
            '<html>T<sub>sm0</sub> = 273.15 + t<sub>sm0</sub></html>',
            '-',
            '<html>T<sub>a</sub> = 273.15 + t<sub>a</sub></html>',
            '-',
            '<html>T<sub>v</sub> = 273.15 + t<sub>v</sub></html>',
            '-',
            '<html>ρ<sub>a</sub> = 353/T<sub>a</sub></html>',
            '<html>ρ<sub>v</sub> = 353/T<sub>v</sub></html>',
        )
        SYMBOLS = (
            '<html>P<sub>sv</sub></html>',
            '<html>t<sub>sm0</sub></html>',
            '<html>T<sub>sm0</sub></html>',
            '<html>t<sub>a</sub></html>',
            '<html>T<sub>a</sub></html>',
            '<html>t<sub>v</sub></html>',
            '<html>T<sub>v</sub></html>',
            'h',
            '<html>ρ<sub>a</sub></html>',
            '<html>ρ<sub>v</sub></html>',
        )
        UNITS = (
            'Па',
            '<sup>o</sup>C',
            'K',
            '<sup>o</sup>C',
            'K',
            '<sup>o</sup>C',
            'K',
            'м',
            'кг/м<sup>3</sup>',
            'кг/м<sup>3</sup>',
        )
        VALUES_TOOLTIPS = {
            0: 'Из таблицы ХОВС',
            1: 'Из расчёта ДУ',
            3: 'Фактическая температура\nпри испытании',
            5: 'Из расчёта ДУ',
            7: 'Из расчёта ДУ или\nпо фактическим измерениям',
        }
        INPUTS_TOOLTIPS = (
            '0...2 500',
            '0...1 000',
            '-50...50',
            '0...1 000',
            '0...300.00',
        )


    class TABLE2:
        NAME = 'table_2'
        ROWS = 5
        COLUMNS = 5
        HEADER = (
            ' Cредняя плотность газа в вытяжном канале (усредненная по значениям температуры Tsm0 и Tv)',
            ' Давление (разряжение) в вытяжном канапе перед вентилятором при температуре перемещаемого воздуха Ta',
            ' Расчетное давление вентилятора при температуре Ta',
            ' Объемный расход перемещаемого вентилятором воздуха при температуре Ta',
            ' Массовый расход перемещаемого вентилятором воздуха при температуре Ta',
        )
        FORMULAS = (
            '<html>ρ<sub>sm</sub> = 2·ρ<sub>a</sub>·T<sub>a</sub>/(T<sub>sm0</sub> + T<sub>v</sub>)</html>',
            '<html>P<sub>sa</sub> = P<sub>sv</sub>·ρ<sub>v</sub>/1.2 + g·h·(ρ<sub>a</sub> - ρ<sub>sm</sub>)</html>',
            '<html>P<sub>sa</sub>·1.2/ρ<sub>v</sub></html>',
            '<html>L<sub>a</sub> = f(P<sub>sa</sub>·1.2/ρ<sub>v</sub>)</html>',
            '<html>G<sub>a</sub> = ρ<sub>a</sub>·L<sub>a</sub>/3600</html>',
        )
        SYMBOLS = (
            '<html>ρ<sub>sm0</sub></html>',
            '<html>P<sub>sa</sub></html>',
            '-',
            '<html>L<sub>a</sub></html>',
            '<html>G<sub>a</sub></html>',
        )
        UNITS = (
            'кг/м<sup>3</sup>',
            'Па',
            'Па',
            'м<sup>3</sup>/ч',
            'кг/с',
        )
        VALUE_TOOLTIP = 'Расход определяется по характеристике\nвентилятора используя давление,\nрассчитанное выше'
        INPUT_TOOLTIP = '0...100 000'


    class DEFAULT_TABLE:
        NAME = 'default_table'
        ROWS = 14
        COLUMNS = 5
        EDITABLE_ROWS = (1, 2, 3, 4, 5, 10, 11)
        SPAN_ROWS = (4, 10)
        HEADER = (
            ' Давление (разряжение) в вытяжном канале у ближайшего к вентилятору закрытого дымоприемного устройства при температуре\n перемещаемого воздуха Та',
            ' Длина вытяжного канала на участке от вентилятора к ближайшему дымоприемному устройству',
            ' Коэффициенты местного сопротивления вытяжного канала на участке от вентилятора к ближайше­му дымоприемному устройству',
            ' Коэффициент эквивалентной шероховатости вытяжного канала',
            ' Размер вытяжного канала на участке от вентилятора к ближайше­му дымоприемному устройству',
            '',
            ' Эквивалентный гидравлический диаметр вытяжного канала на участке от вентилятора к ближайшему дымоприемному устройству',
            ' Площадь проходного сечения вытяжного канала на участке от вентилятора к ближайше­му дымоприемному устройству',
            ' Коэффициент сопротивления трения вытяжного канала на участке от вентилятора к ближайше­му дымоприемному устройству',
            ' Подсосы воздуха через ближайшие к вентилятору дымоприемные устройства (противопожарные нормально закрытые клапаны)',
            ' Размер дымоприемного устройства (клапана) ближайшего к вентилятору',
            '',
            ' Площадь проходного сечения дымоприемного устройства (клапана) ближайшего к вентилятору',
            ' Массовый расход перемещаемого в вытяжном канале воздуха у закрытого дымоприемного устройства на рассматриваемом участке',
        )
        FORMULAS_0 = (
            'P<sub>sn</sub> = P<sub>sa</sub> - 0.5·ρ<sub>a</sub>·(∑ζ<sub>n</sub> + λ<sub>n</sub>·l<sub>n</sub>/d<sub>en</sub>)·(G<sub>a</sub>/(ρ<sub>a</sub>·F<sub>n</sub>))<sup>2</sup>',
            '-',
            '-',
            '-',
            '-',
            '-',
            'd<sub>en</sub> = 2·a<sub>n</sub>·b<sub>n</sub>/(a<sub>n</sub> + b<sub>n</sub>)',
            'F<sub>n</sub> = a<sub>n</sub>·b<sub>n</sub>',
            '-',
            'ΔG<sub>pn</sub> = F<sub>dpn</sub>·(P<sub>sn</sub>/S<sub>dpn</sub>)<sup>1/2</sub>',
            '-',
            '-',
            'F<sub>dpn</sub> = a<sub>dpn</sub>·b<sub>dpn</sub>',
            'G<sub>1</sub> = G<sub>a</sub> - ΔG<sub>dpn</sub>',
        )
        FORMULAS_N = (
            'P<sub>s%d</sub> = P<sub>s%s</sub> - 0.5·ρ<sub>a</sub>·(∑ζ<sub>%d</sub> + λ<sub>%d</sub>·l<sub>%d</sub>/d<sub>e%d</sub>)·(G<sub>a</sub>/(ρ<sub>a</sub>·F<sub>%d</sub>))<sup>2</sup>',
            '-',
            '-',
            '-',
            '-',
            '-',
            'd<sub>e%d</sub> = 2·a<sub>%d</sub>·b<sub>%d</sub>/(a<sub>%d</sub> + b<sub>%d</sub>)',
            'F<sub>%d</sub> = a<sub>%d</sub>·b<sub>%d</sub>',
            '-',
            'ΔG<sub>p%d</sub> = F<sub>dp%d</sub>·(P<sub>s%d</sub>/S<sub>dp%d</sub>)<sup>1/2</sub>',
            '-',
            '-',
            'F<sub>dp%d</sub> = a<sub>dp%d</sub>·b<sub>dp%d</sub>',
            'G<sub>%d</sub> = G<sub>%d</sub> - ΔG<sub>dp%d</sub>',
        )
        SYMBOLS_0 = (
            'P<sub>sn</sub>',
            'l<sub>n</sub>',
            'ζ<sub>n</sub>',
            'k<sub>эn</sub>',
            'a<sub>n</sub>',
            'b<sub>n</sub>',
            'd<sub>en</sub>',
            'F<sub>n</sub>',
            'λ<sub>n</sub>',
            'ΔG<sub>dpn</sub>',
            'a<sub>dpn</sub>',
            'b<sub>dpn</sub>',
            'F<sub>dpn</sub>',
            'G<sub>1</sub>',
        )
        SYMBOLS_N = (
            'P<sub>s%d</sub>',
            'l<sub>%d</sub>',
            'ζ<sub>%d</sub>',
            'k<sub>э%d</sub>',
            'a<sub>%d</sub>',
            'b<sub>%d</sub>',
            'd<sub>e%d</sub>',
            'F<sub>%d</sub>',
            'λ<sub>%d</sub>',
            'ΔG<sub>dp%d</sub>',
            'a<sub>dp%d</sub>',
            'b<sub>dp%d</sub>',
            'F<sub>dp%d</sub>',
            'G<sub>%d</sub>',
        )
        UNITS = (
            'Па',
            'м',
            '-',
            '-',
            'м',
            'м',
            'м',
            'м<sup>2</sup>',
            '-',
            'кг/с',
            'м',
            'м',
            'м<sup>2</sup>',
            'кг/с',
        )
        INPUTS_TOOLTIPS = (
            '0...50.00',
            '''<html>
            0.1_________Листовая сталь (новая)<br>
            0.15________Листовая сталь (б/у)<br>
            0.015-0.06__Алюминий<br>
            0.3-0.8_____Бетон с затиркой<br>
            2.5_________ЖБ<br>
            1___________Шлакогипс<br>
            1.5-2_______Шлакобетон<br>
            5-10________Кладка (каналы в стене)<br>
            0.5-3_______Тоже оштук. ЦПС<br>
            10-15_______Тоже оштук. по сетке<br>
            0.11________Асбест
            </html>''',
            '0.1...5.0',
            '0.1...6.0',
            '0.1...5.0',
            '0.1...6.0',
        )
        EXPORT_ROWS = (0, 9, 13)


    class BOARD:
        LABELS = (
            ' Объемный расход воздуха, поступающего через открытое дымоприемное устройство при температуре Та',
            '',
            'L<sub>0</sub> = 3600·G<sub>2</sub>/ρ<sub>a</sub>',
            '',
            'м<sup>3</sup>/ч',
            '',
        )
        NUM_TABLES_TOOLTIP = 'Количество\nучастков'


    class SAVE_OPEN:
        TABLE1 = (0, 1, 3, 5, 7)
        TABLE2 = 3
        DEFAULT_TABLE = (1, 2, 3, 4, 5, 10, 11)


    class EXPORT:
        TITLE = 'Определение расчетного значения требуемого расхода воздуха через открытые дымоприемные устройства в приемо-сдаточных и периодических испытаниях противодымной вентиляции'
        HEADER = (
            'Наименование расчетной величины',
            'Обозначение',
            'Величина',
            'Ед.изм.',
        )

        class TABLE1:
            HEADER = (
                'Приведенное статическое давление вентилятора',
                'Установленная при проектировании испытуемой системы вытяжной противодымной вентиляции значение температуры продуктов горения, непосредственно удаляемых из защищаемого помещения (коридора)',
                '',
                'Температура воздуха в помещениях и в вытяжном вентиляционном канале при проведении приемо-сдаточных и периодических испытаний',
                '',
                'Установленная при проектировании испытуемой системы вытяжной противодымной вентиляции значение температуры продуктов горения, перемещаемых вентилятором',
                '',
                'Разность уровней фактического расположения входного устройства вентилятора и открытого дымоприемного устройства вытяжного канала',
                'Плотность воздуха при температуре Ta',
                'Плотность удаляемого газа, перемещаемого вентилятором (при температуре Tv)',
            )
            SYMBOLS = (
                'Psv',
                'tsm0',
                'Tsm0',
                'ta',
                'Ta',
                'tv',
                'Tv',
                'h',
                'ρa',
                'ρv',
            )
            UNITS = (
                'Па',
                '℃',
                'K',
                '℃',
                'K',
                '℃',
                'K',
                'м',
                'кг/м3',
                'кг/м3',
            )


        class TABLE2:
            HEADER = (
                'Cредняя плотность газа в вытяжном канале (усредненная по значениям температуры Tsm0 и Tv)',
                'Давление (разряжение) в вытяжном канапе перед вентилятором при температуре перемещаемого воздуха Ta',
                'Расчетное давление вентилятора при температуре Ta',
                'Объемный расход перемещаемого вентилятором воздуха при температуре Ta',
                'Массовый расход перемещаемого вентилятором воздуха при температуре Ta',
            )
            FORMULAS = (
                'ρsm = 2·ρa·Ta/(Tsm0 + Tv)',
                'Psa = Psv·ρv/1.2 + g·h·(ρa - ρsm)',
                'Psa·1.2/ρv',
                'La = f(Psa·1.2/ρv)',
                'Ga = ρa·La/3600',
            )
            SYMBOLS = (
                'ρsm0',
                'Psa',
                '-',
                'La',
                'Ga',
            )
            UNITS = (
                'кг/м3',
                'Па',
                'Па',
                'м3/ч',
                'кг/с',
            )


        class DEFAULT_TABLE:
            HEADER = (
                'Давление (разряжение) в вытяжном канале у ближайшего к вентилятору закрытого дымоприемного устройства при температуре перемещаемого воздуха Та',
                'Подсосы воздуха через ближайшие к вентилятору дымоприемные устройства (противопожарные нормально закрытые клапаны)',
                'Массовый расход перемещаемого в вытяжном канале воздуха у закрытого дымоприемного устройства на рассматриваемом участке',
            )
            SYMBOLS_0 = (
                'Psn',
                'ΔGdpn',
                'G1',
            )
            SYMBOLS_N = (
                'Ps%d',
                'ΔGdp%d',
                'G%d',
            )
            UNITS = (
                'Па',
                'кг/с',
                'кг/с',
            )


        class RESULT:
            TITLE = 'Результат расчёта'
            DATA = (
                'Объемный расход воздуха, поступающего через открытое дымоприемное устройство при температуре Та',
                'L0',
                '',
                'м3/ч',
            )




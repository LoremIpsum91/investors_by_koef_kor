import openpyxl

letters = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']


def excel(letter):
    values = []
    wb = openpyxl.load_workbook('data.xlsx')
    sheet = wb['A1']

    for i in range(2, 1012):
        value = sheet['%s%d' % (letter, i)].value
        values.append(value)

    return values


h = 1

dates = excel('A')


def koef_kor(go):
    # m, k = input('Введите границы через пробел (учтите то, что вторая дата не включается в промежуток)').split()
    m, k = go.split()
    all_kor = []
    global h, dates, s
    while (dates[-h].split('.')[1]) + '.' + dates[-h].split('.')[2] != str(m):
        h += 1
    for i in range(len(letters)):
        first = excel(letters[i])
        for j in range(i+1, len(letters)):
            second = excel(letters[j])
            only_first = []
            only_second = []
            s = h
            while (dates[-s].split('.')[1] + '.' + dates[-s].split('.')[2]) != k:
                only_first.append(float(first[-s]))
                only_second.append(float(second[-s]))
                s += 1
            # if float(first[-h]) >= float(first[-s]) or float(second[-h]) >= float(second[-s]):
            #     continue
            # else:
            # print(only_first, only_second)
            sredn_znach1, sredn_znach2 = sum(only_first) / len(only_first), sum(only_second) / len(only_second)
            otklonenie1, otklonenie2 = 0, 0
            for g in range(len(only_first)):
                otklonenie1 += (float(only_first[g]) - sredn_znach1) ** 2
                otklonenie2 += (float(only_second[g]) - sredn_znach2) ** 2
            srednee_kvad_otkl1 = (otklonenie1 / (len(only_first))) ** (1 / 2)
            srednee_kvad_otkl2 = (otklonenie2 / (len(only_second))) ** (1 / 2)
            umnozh = 0
            for l in range(len(only_first)):
                umnozh += (float(only_first[l]) * float(only_second[l]))
            # print(srednee_kvad_otkl1, srednee_kvad_otkl2)
            kor = ((umnozh / len(only_first)) - (sredn_znach1 * sredn_znach2)) / (
                srednee_kvad_otkl1 * srednee_kvad_otkl2)
            all_kor.append([kor, i, j])
    all_kor = sorted(all_kor)
    return all_kor


# def test():
#     first = [7.1, 6.3, 6.2, 5.8, 7.7, 6.8, 6.7, 5.9, 5.7, 5.1]
#     second = [42, 107, 100, 60, 78, 79, 90, 54]
#     only_first = []
#     only_second = []
#     sredn_znach1 = sum(first) / len(first)
#     otklonenie1, otklonenie2 = 0, 0
#     for g in range(len(first)):
#         otklonenie1 += (float(first[g]) - sredn_znach1) ** 2
#     srednee_kvad_otkl1 = ((otklonenie1 / (len(first)))) ** (1/2)
#     print(srednee_kvad_otkl1, sredn_znach1, otklonenie1)
#     return


def start():
    n = 3000000
    boris = 3000000
    x = n/6

    kol_vo1 = 0
    kol_vo2 = 0
    kol_vo3 = 0
    kol_vo4 = 0
    kol_vo5 = 0
    kol_vo6 = 0

    kol_vo11 = 0
    kol_vo21 = 0
    kol_vo31 = 0
    kol_vo41 = 0
    kol_vo51 = 0
    kol_vo61 = 0

    kor = koef_kor('01.%d 07.%d' % (2016, 2016))
    kor_n = kor.copy()
    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0


    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('07.%d 01.%d' % (2016, 2017))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('01.%d 07.%d' % (2017, 2017))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('07.%d 01.%d' % (2017, 2018))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('01.%d 07.%d' % (2018, 2018))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('07.%d 01.%d' % (2018, 2019))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('01.%d 07.%d' % (2019, 2019))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()

    kol_vo1 = (n / float(excel(letters[kor[0][1]])[-h]))
    kol_vo2 = (n / float(excel(letters[kor[0][2]])[-h]))
    kol_vo3 = (n / float(excel(letters[kor[1][1]])[-h]))
    kol_vo4 = (n / float(excel(letters[kor[1][2]])[-h]))
    kol_vo5 = (n / float(excel(letters[kor[2][1]])[-h]))
    kol_vo6 = (n / float(excel(letters[kor[2][2]])[-h]))
    n = 0

    kol_vo11 = (boris / float(excel(letters[kor[-1][1]])[-h]))
    kol_vo21 = (boris / float(excel(letters[kor[-1][2]])[-h]))
    kol_vo31 = (boris / float(excel(letters[kor[-2][1]])[-h]))
    kol_vo41 = (boris / float(excel(letters[kor[-2][2]])[-h]))
    kol_vo51 = (boris / float(excel(letters[kor[-3][1]])[-h]))
    kol_vo61 = (boris / float(excel(letters[kor[-3][2]])[-h]))
    boris = 0
    print(n, boris)
    kor = koef_kor('07.%d 12.%d' % (2019, 2019))

    n += float(excel(letters[kor_n[0][1]])[-h]) * kol_vo1
    n += float(excel(letters[kor_n[0][2]])[-h]) * kol_vo2
    n += float(excel(letters[kor_n[1][1]])[-h]) * kol_vo3
    n += float(excel(letters[kor_n[1][2]])[-h]) * kol_vo4
    n += float(excel(letters[kor_n[2][1]])[-h]) * kol_vo5
    n += float(excel(letters[kor_n[2][2]])[-h]) * kol_vo6

    boris += float(excel(letters[kor_n[-1][1]])[-h]) * kol_vo11
    boris += float(excel(letters[kor_n[-1][2]])[-h]) * kol_vo21
    boris += float(excel(letters[kor_n[-2][1]])[-h]) * kol_vo31
    boris += float(excel(letters[kor_n[-2][2]])[-h]) * kol_vo41
    boris += float(excel(letters[kor_n[-3][1]])[-h]) * kol_vo51
    boris += float(excel(letters[kor_n[-3][2]])[-h]) * kol_vo61
    print(n, boris)
    kor_n = kor.copy()


    return n, boris




print(start())
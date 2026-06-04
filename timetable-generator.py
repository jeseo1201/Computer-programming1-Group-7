"""
시간표 자동 생성기 (2026학년도 1학기) - 17학점 맞춤형
실행 방법: 이 .py 파일과 subjectTable.csv를 같은 폴더에 두고 실행
"""

import os
import random
import pandas as pd

SCIENCE_OPTIONS = {
    '1': ('일반 물리학1', '일반물리학및실험1'),
    '2': ('화학1', '일반화학및실험1'),
    '3': ('생물학1', '일반생물학1'),
    '4': ('일반수학1및연습', '일반수학1및연습'),
    '5': ('기타 교양과목', None),
}

def clean(value):
    if pd.isna(value):
        return ''
    return str(value).strip()

def parseCredit(value):
    try:
        return int(str(value).split('/')[-1])
    except Exception:
        return 0

class Subject:
    def __init__(self, row):
        self.name = clean(row.get('교과목명', ''))
        self.section = clean(row.get('분반', ''))
        self.credit = parseCredit(row.get('시간/학점', '0/0'))
        self.time = clean(row.get('요일및교시(강의실)', ''))
        self.prof = clean(row.get('담당교수', ''))
        self.type = clean(row.get('1전공기준 이수구분', ''))

def parseSlots(timeText):
    slots = set()
    if clean(timeText) == '':
        return slots

    for part in str(timeText).split(','):
        part = part.strip().split('(')[0]
        if len(part) < 2:
            continue

        day = part[0]
        timePart = part[1:]

        try:
            if '~' in timePart:
                start, end = map(int, timePart.split('~'))
                for period in range(start, end + 1):
                    slots.add((day, period))
            else:
                slots.add((day, int(timePart)))
        except Exception:
            pass

    return slots

def rowToInfo(row):
    return Subject(row)

def uniqueNames(rows):
    result = []
    for row in rows:
        name = clean(row.get('교과목명', ''))
        if name and name not in result:
            result.append(name)
    return result

def findCombos(pools, targetCredit=17, maxResults=5):
    results = []

    def backtrack(idx, chosen, usedSlots, usedNames):
        if len(results) >= maxResults:
            return

        if idx == len(pools):
            total = sum(item.credit for item in chosen)
            if total == targetCredit:
                results.append(chosen[:])
            return

        pool = pools[idx][:]
        random.shuffle(pool)

        for subject in pool:
            if subject.name in usedNames:
                continue

            slots = parseSlots(subject.time)
            if usedSlots & slots:
                continue

            chosen.append(subject)
            usedNames.add(subject.name)
            backtrack(idx + 1, chosen, usedSlots | slots, usedNames)
            chosen.pop()
            usedNames.remove(subject.name)

    backtrack(0, [], set(), set())
    return results

def displayLength(text):
    total = 0
    for ch in str(text):
        if ord(ch) > 127:
            total += 2
        else:
            total += 1
    return total

def fitCenter(text, width):
    text = str(text)
    result = ''
    length = 0

    for ch in text:
        chLength = 2 if ord(ch) > 127 else 1
        if length + chLength > width - 1:
            result += '…'
            break
        result += ch
        length += chLength

    padding = max(0, width - displayLength(result))
    left = padding // 2
    right = padding - left
    return ' ' * left + result + ' ' * right

def printTimetable(subjects):
    days = ['월', '화', '수', '목', '금', '토', '일']
    usedDays = set()
    grid = {}

    for subject in subjects:
        for part in subject.time.split(','):
            part = part.strip().split('(')[0]
            if len(part) < 2:
                continue

            day = part[0]
            timePart = part[1:]

            if day not in days:
                continue

            try:
                if '~' in timePart:
                    start, end = map(int, timePart.split('~'))
                    for period in range(start, end + 1):
                        grid[(day, period)] = subject.name
                        usedDays.add(day)
                else:
                    period = int(timePart)
                    grid[(day, period)] = subject.name
                    usedDays.add(day)
            except Exception:
                pass

    if not grid:
        print('\n  시간표 정보 없음')
        return

    usedDays = [day for day in days if day in usedDays]
    periods = sorted(set(period for _, period in grid.keys()))

    timeWidth = 8
    dayWidth = 14

    def border(left, middle, right):
        line = left + '─' * timeWidth
        for _ in usedDays:
            line += middle + '─' * dayWidth
        return '  ' + line + right

    print('\n' + border('┌', '┬', '┐'))

    header = '  │' + fitCenter('교시', timeWidth)
    for day in usedDays:
        header += '│' + fitCenter(day, dayWidth)
    print(header + '│')

    print(border('├', '┼', '┤'))

    for period in range(periods[0], periods[-1] + 1):
        row = '  │' + fitCenter(f'{period}교시', timeWidth)
        for day in usedDays:
            row += '│' + fitCenter(grid.get((day, period), ''), dayWidth)
        print(row + '│')

        if period != periods[-1]:
            print(border('├', '┼', '┤'))

    print(border('└', '┴', '┘'))

def printResult(results):
    print(f'\n탐색 완료: 추천 시간표 {len(results)}개 발견')

    if not results:
        print('조건에 맞는 시간표 조합이 없습니다.')
        return

    for idx, result in enumerate(results, 1):
        total = sum(subject.credit for subject in result)
        print(f'\n[추천 시간표 #{idx}] 총 {total}학점')
        printTimetable(result)

        print('\n  과목 요약:')
        for subject in result:
            print(
                f"  - [{subject.type}] {subject.name} "
                f"| {subject.section}분반 | {subject.credit}학점 | {subject.prof}"
            )

def run():
    folder = os.path.dirname(os.path.abspath(__file__))
    csvPath = os.path.join(folder, 'subjectTable.csv')

    if not os.path.exists(csvPath):
        print(f'파일을 찾을 수 없습니다: {csvPath}')
        return

    df = pd.read_csv(csvPath, encoding='utf-8-sig', dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all').reset_index(drop=True)

    targetMajor = input('본인의 전공(학과명 정확히): ').strip()
    offDay = input('공강 희망 요일 (없으면 Enter): ').strip()

    rows = []
    for _, row in df.iterrows():
        remark = clean(row.get('비고', ''))
        timeText = clean(row.get('요일및교시(강의실)', ''))

        if remark and targetMajor not in [x.strip() for x in remark.split(',')]:
            continue
        if offDay and offDay in timeText:
            continue

        rows.append(row.to_dict())

    filtered = pd.DataFrame(rows, columns=df.columns)
    nameCol = filtered['교과목명'].fillna('').str.strip()
    typeCol = filtered['1전공기준 이수구분'].fillna('').str.strip()
    deptCol = filtered['개설전공(과)'].fillna('').str.strip()

    peopleRows = filtered[nameCol == '인간학1'].to_dict('records')
    designRows = filtered[nameCol == 'I-DESIGN'].to_dict('records')

    coreRows = filtered[
        (typeCol == '중핵교양필수') &
        (nameCol != 'I-DESIGN')
    ].to_dict('records')

    majorRows = filtered[
        (typeCol == '전공기초') &
        (deptCol == targetMajor)
    ].to_dict('records')

    liberalRows = filtered[typeCol == '자유선택교양'].to_dict('records')

    print('\n[추가 중핵교양필수 1과목 선택]')
    print('1) 랜덤 배정')
    print('2) 영역 지정 (인문사회 / 자연과학 / 휴먼테크 / 글로벌외국어)')
    print('3) 과목명 직접 검색')
    coreChoice = input('방식을 번호로 선택하세요: ').strip()

    if coreChoice == '2':
        domain = input('원하는 영역을 정확히 입력하세요: ').strip()
        corePool = [r for r in coreRows if clean(r.get('중핵 편성영역', '')) == domain]
        if not corePool:
            print('해당 영역 과목이 없어 전체 중핵교양에서 찾습니다.')
            corePool = coreRows
    elif coreChoice == '3':
        keyword = input('검색할 과목명을 입력하세요: ').strip()
        corePool = [r for r in coreRows if keyword in clean(r.get('교과목명', ''))]
        if not corePool:
            print('해당 과목이 없어 전체 중핵교양에서 찾습니다.')
            corePool = coreRows
    else:
        corePool = coreRows

    majorNames = uniqueNames(majorRows)
    if not majorNames:
        print(f'[{targetMajor}] 전공기초 과목을 찾지 못했습니다.')
        return

    print('\n[전공기초 과목 목록]')
    print(', '.join(majorNames))

    while True:
        pickedMajor = input('듣고 싶은 전공기초 과목명 1개: ').strip()
        if pickedMajor in majorNames:
            break
        print('목록에 있는 과목명을 정확히 입력해주세요.')

    print('\n[추가 과목 선택: 아래 중 2개]')
    for key, value in SCIENCE_OPTIONS.items():
        print(f'{key}) {value[0]}')

    while True:
        extraInput = input('번호 2개 입력 예시 1,4: ').strip()
        pickedOptions = [x.strip() for x in extraInput.split(',')]

        if len(pickedOptions) == 2 and len(set(pickedOptions)) == 2:
            if all(option in SCIENCE_OPTIONS for option in pickedOptions):
                break

        print('1, 2, 3, 4 중 서로 다른 번호 2개를 입력해주세요.')

    pools = []

    if peopleRows:
        pools.append([rowToInfo(r) for r in peopleRows])
    if designRows:
        pools.append([rowToInfo(r) for r in designRows])
    if corePool:
        pools.append([rowToInfo(r) for r in corePool])

    pickedMajorRows = [r for r in majorRows if clean(r.get('교과목명', '')) == pickedMajor]
    pools.append([rowToInfo(r) for r in pickedMajorRows])

    for option in pickedOptions:
        subjectName = SCIENCE_OPTIONS[option][1]

        if subjectName is None:
            extraRows = liberalRows
        else:
            extraRows = filtered[nameCol == subjectName].to_dict('records')

        pools.append([rowToInfo(r) for r in extraRows])

    print('\n시간표 조합 탐색 중...')
    results = findCombos(pools, targetCredit=17, maxResults=5)
    printResult(results)

if __name__ == '__main__':
    run()


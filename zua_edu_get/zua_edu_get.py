import requests, re, time
from hashlib import sha1

url = "http://jwglxt.zua.edu.cn/eams/courseTableForStd!courseTable.action"
login_url = "http://jwglxt.zua.edu.cn/eams/loginExt.action"

# 网页实例
s = requests.Session()

username = input("学号: ")
password = input("密码: ")

#----------登录页面
page = s.get(login_url).text
time.sleep(2)

#密钥
m = re.search(r"SHA1\('([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-)'", page)
sha1p = sha1((m.group(1) + password).encode()).hexdigest()

#----------登录
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0",
    "Referer": "http://jwglxt.zua.edu.cn/eams/loginExt.action",
    "Origin": "http://jwglxt.zua.edu.cn",
    "Content-Type": "application/x-www-form-urlencoded",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Upgrade-Insecure-Requests": "1",
    }

data = {
    "needModify": "0",
    "username": username,
    "password": sha1p,
    "session_locale": "zh_CN",
    "login_from": "login_from"
    }

s.post(login_url, headers=headers, data=data)
time.sleep(2)

#----------获取课表
#模拟点击访问
r = s.get("http://jwglxt.zua.edu.cn/eams/courseTableForStd.action")
time.sleep(1)

#查询时间id
tag_id = "semesterBar" + re.search(r'semesterBar(\d+)Semester', r.text).group(1) + "Semester"
cur = s.cookies.get("semester.id") or "297"
r2 = s.post(url = "http://jwglxt.zua.edu.cn/eams/dataQuery.action", data = {
    "tagId": tag_id,
    "dataType": "semesterCalendar",
    "value": cur,
    "empty": "false"
    })
beg = input("输入学期年份后两位（如2026-2027的输入26即可）：")
sx = input("输入学期（1上学期，2下学期）：")
tcode = re.search(rf'id:(\d+),schoolYear:"{"20" + beg + "-20" + str(int(beg) + 1)}",name:"{sx}"', r2.text).group(1)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0",
    "Referer": "http://jwglxt.zua.edu.cn/eams/courseTableForStd.action",
    "Origin": "http://jwglxt.zua.edu.cn",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Accept": "*/*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "x-requested-with": "XMLHttpRequest"
    }

ids = re.search(r'"ids","(\d+)"', r.text).group(1)

data = {
    "ignoreHead": "1",
    "setting.kind": "std",
    "startWeek": "",
    "semester.id": tcode,
    "ids": ids
    }

#获取课表
resp = s.post(url = url, headers = headers, data = data)

#print(resp.text)

# ---------- 解析课表并生成 Excel（追加部分） ----------
import openpyxl, os

def split_args(s):
    out, buf, depth, q = [], "", 0, None
    for c in s:
        if q:
            buf += c
            if c == q:
                q = None
        elif c in '"\'':
            buf += c
            q = c
        elif c == '(':
            depth += 1
            buf += c
        elif c == ')':
            depth -= 1
            buf += c
        elif c == ',' and depth == 0:
            out.append(buf.strip())
            buf = ""
        else:
            buf += c
    if buf.strip():
        out.append(buf.strip())
    return out

pat = re.compile(r"new TaskActivity\((.*?)\);\s*index =(\d+)\*unitCount\+(\d+);", re.S)
grid = [[None] * 7 for _ in range(10)]  # grid[节次][星期几]

for mm in pat.finditer(resp.text):
    args = split_args(mm.group(1))
    if len(args) < 7:
        continue
    course = args[3].strip('"').split("(")[0] or args[2].strip('"').split("(")[0]
    room = args[5].strip('"') or args[4].strip('"')
    bits = args[6].strip('"')
    weeks = [str(i + 1) for i, ch in enumerate(bits) if ch == "1"]
    day = int(mm.group(2))
    unit = int(mm.group(3))
    if day >= 7 or unit >= 10:
        continue
    seg = resp.text[:mm.start()]
    tpos = [t.start() for t in re.finditer(r"var teachers = \[", seg)]
    if tpos:
        names = re.findall(r'name:"([^"]+)"', seg[tpos[-1]:])
        teacher = "、".join(dict.fromkeys(names)) or "未知"
    else:
        teacher = "未知"
    ent = (course, room, teacher, ",".join(weeks))
    if grid[unit][day] is None:
        grid[unit][day] = [ent]
    else:
        for i, e in enumerate(grid[unit][day]):
            if e[0] == course and e[1] == room and e[2] == teacher:
                merged = sorted(set(e[3].split(",")) | set(weeks), key=int)
                grid[unit][day][i] = (e[0], e[1], e[2], ",".join(merged))
                break
        else:
            grid[unit][day].append(ent)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "课表"
days = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
al = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)

ws.cell(1, 1, "节次/周次")
for d, txt in enumerate(days):
    ws.cell(1, 2 + d, txt)

def same_block(a, b):
    if not a or not b or len(a) != len(b):
        return False
    return all(x[0] == y[0] and x[1] == y[1] and x[2] == y[2] for x, y in zip(a, b))

for d in range(7):
    u = 0
    while u < 10:
        if grid[u][d] is None:
            u += 1
            continue
        span = 1
        while u + span < 10 and same_block(grid[u][d], grid[u + span][d]):
            span += 1
        lines = []
        for course, room, teacher, w in grid[u][d]:
            lines += [course, room, teacher, w]
        val = "\n".join(lines)
        r, c = 2 + u, 2 + d
        if span > 1:
            ws.merge_cells(start_row=r, start_column=c, end_row=r + span - 1, end_column=c)
        ws.cell(r, c, val).alignment = al
        u += span

for u in range(10):
    ws.cell(2 + u, 1, "第" + str(u + 1) + "节").alignment = al
for i, wd in enumerate([10] + [17] * 7, start=1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd

out = os.path.join(os.path.expanduser("~"), "Desktop", "课表.xlsx")
wb.save(out)
print("已生成课表：", out)
os.system("pause")
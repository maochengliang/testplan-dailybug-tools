import asyncio
import json

from django.shortcuts import render, HttpResponse, redirect
from django.contrib import messages
# from starlette.responses import FileResponse
from django.http import FileResponse
from urllib.parse import quote
from datetime import datetime, timedelta
import xlrd
import xlwt
import requests
import re
from xlwt import XFStyle, Pattern
from dateutil import rrule
from xlutils.copy import copy
import calendar


# Create your views here.
def hello(request):
    # return HttpResponse("hello world")
    return render(request, "base.html")


def base_two(request):
    # return HttpResponse("hello world")
    return render(request, "base_two.html")


def first_page(request):
    return render(request, "extend.html")


def create_html(request):
    if request.method == "GET":
        messages.get_messages(request)
        context = {
            'messages': messages.get_messages(request),
        }
        return render(request, 'plan.html', context)
        # return render(request, "plan.html")

    print(request.POST)
    # 计划标题
    title = request.POST.get("title")
    # 测试目标
    test_target = request.POST.get("test_target")

    # 计划需要新建的设备端口
    port = request.POST.getlist("port")

    # =============测试范围==============================================
    # 业务名称集
    p_name = request.POST.getlist("p_name")

    # 页面/功能集
    page = request.POST.getlist("page")

    # 模块集
    model = request.POST.getlist("model")

    # 负责人集
    people = request.POST.getlist("people")

    # ===================测试进度===============================

    # 用例时长
    case_time = request.POST.get("case_time")

    # 用例实际时长
    case_shiji_time = request.POST.get("case_shiji_time")

    # 评审日期
    pingshen_time = request.POST.get("pingshen_time")

    # 实际评审日期
    pingshen_shiji_time = request.POST.get("pingshen_shiji_time")

    # 提测日期
    tice = request.POST.get("tice_time")

    # 实际提测日期
    tice_shiji_time = request.POST.get("tice_shiji_time")

    # 一轮测试时长
    yilun_time = request.POST.get("yilun_time")

    # 实际一轮测试时长
    yilun_shiji_time = request.POST.get("yilun_shiji_time")

    # 验收时间
    yanshou_time = request.POST.get("yanshou_time")

    # 实际验收时间
    yanshou_shiji_time = request.POST.get("yanshou_shiji_time")

    # 上线时间
    shangxian_time = request.POST.get("shangxian_time")

    # 实际上线时间
    shangxian_shiji_time = request.POST.get("shangxian_shiji_time")

    # =================参考文档==================================
    # 需求链接
    xuqiu_link = request.POST.get("xuqiu_link")

    # 设计稿链接
    sheji_link = request.POST.get("sheji_link")

    # 接口文档链接
    api_link = request.POST.get("api_link")

    # =================测试策略=====================================

    # 功能测试执行结果
    gongneng_result = request.POST.get("gongneng_result")

    # 接口测试执行结果
    api_result = request.POST.get("api_result")

    # 界面测试执行结果
    jiemian_result = request.POST.get("jiemian_result")

    author = request.POST.get("author")

    project = request.POST.get("project")

    xuqiu_id = request.POST.get("xuqiu_id")

    startTime = request.POST.get("start-time")
    endTime = request.POST.get("end-time")

    # 组装数据
    obj = {
        "test_target": test_target,
        "case_time": case_time,
        "case_shiji_time": case_shiji_time,
        "pingshen_time": pingshen_time,
        "pingshen_shiji_time": pingshen_shiji_time,
        "tice_time": tice,
        "tice_shiji_time": tice_shiji_time,
        "yilun_time": yilun_time,
        "yilun_shiji_time": yilun_shiji_time,
        "yanshou_time": yanshou_time,
        "yanshou_shiji_time": yanshou_shiji_time,
        "shangxian_time": shangxian_time,
        "shangxian_shiji_time": shangxian_shiji_time,
        "xuqiu_link": xuqiu_link,
        "sheji_link": sheji_link,
        "api_link": api_link,
        "gongneng_result": gongneng_result,
        "api_result": api_result,
        "jiemian_result": jiemian_result
    }

    max_list = [tuple(x) for x in zip(p_name, page, model, people)]
    print("max_list===============", max_list)

    device_list = [None] * len(port)

    for i, device in enumerate(port):
        if device == "1":
            device_list[i] = "PC"
        elif device == "2":
            device_list[i] = "Touch"
        elif device == "3":
            device_list[i] = "Android"
        else:
            device_list[i] = "iOS"

    p = ""
    if project == "新建案":
        p += "1163835346001001629"
    elif project == "中古社区":
        p += "1163835346001001626"
    elif project == "租售":
        p += "1163835346001001637"
    elif project == "工程":
        p += "1163835346001001638"
    elif project == "日常BUG":
        p += "1163835346001001639"

    for dl in device_list:
        new_title = dl + "--" + title
        s = ""
        for l in range(len(max_list)):
            yewustr = f'<tr><td width="150"><p class="ne-p"><span class="ne-text">&nbsp;{max_list[l][0]}</span></p></td><td width="150"><p class="ne-p"><span class="ne-text">&nbsp;{max_list[l][1]}</span></p></td><td width="150"><p class="ne-p"><span class="ne-text">&nbsp;{max_list[l][2]}</span></p></td><td width="150"><p class="ne-p"><span class="ne-text">&nbsp;{dl}</span></p></td><td width="150"><p class="ne-p"><span class="ne-text">&nbsp;{max_list[l][3]}</span></p></td></tr>'
            s += yewustr

        createPlanRequest(obj=obj, yewustr=s, title=new_title, author=author, project=p, xuqiu_id=xuqiu_id, startTime=startTime, endTime=endTime)

        print(s)

    # return redirect("https://www.tapd.cn/63835346/sparrow/test_plan/plan_list")
    messages.success(request, '新建成功！')
    return redirect(request.path)


async def download_html(request):
    if request.method == "GET":
        return render(request, 'utils.html')

    print(type(request.body.decode('utf-8')))
    print(request.body.decode('utf-8'))
    body_param = json.loads(request.body.decode('utf-8'))

    print(type(body_param))
    sTime = body_param["startTime"]
    eTime = body_param["endTime"]
    export_excel(sTime, eTime)
    # 处理完毕文件以后，生成了文件路径
    fileName = "-".join(sTime.split("-")[:2])
    fileName = fileName.split("-")[1]
    filename = f"{fileName}日常Bug统计数据.xls"
    # time.sleep(10)
    await asyncio.sleep(3)
    return FileResponse(open(filename, 'rb'), as_attachment=True)
    # encoded_filename = quote(filename)  # 对文件名进行 URL 编码
    # return FileResponse(
    #     filename,  # 打开文件
    #     filename=encoded_filename,
    #     headers={
    #         'Content-Disposition': f'attachment; filename="{encoded_filename}"'
    #     },
    #     media_type='application/octet-stream',
    # )


def export_excel(startTime, endtTime):
    # 拿到开始时间
    # startTime = startDateEdit.dateTime().toString("yyyy-MM-dd")
    # # 拿到结束时间
    # endtTime = endDateEdit.dateTime().toString("yyyy-MM-dd")

    # 转换字符串截取成  年-月 格式

    query_tapd_bug_write_excel(startTime, startTime, endtTime)

    fileName = "-".join(startTime.split("-")[:2])
    fileName = fileName.split("-")[1]
    print("fileNmae===", fileName)
    readTapdList(fileName)

    writeCount(fileName)


def query_tapd_bug_write_excel(startTime, fileName, endTime):
    start_cookies = {
        '__root_domain_v': '.tapd.cn',
        '_qddaz': 'QD.271048197899749',
        '63835346bug_create_template': '1163835346001000268',
        'tcase_selected_category_id': '1163835346001001771',
        'new_worktable': 'search_filter',
        '_t_uid': '2039622907',
        '_t_crop': '22912571',
        'tapd_div': '101_3',
        'tapdsession': '1678441751d083b6e27a676b7bd3b8642229db7a7fdbf73939741830f8f3f02ef7f75c46f7',
        't_u': '226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8ed6487938d984910ceb15bdf4a7b3bf01dc65e068b228a5fde9333e46ff239bffd38a2bfa9701f215295d673b8fabff94%7C1',
        't_cloud_login': 'myy%40addcn.com',
        'dsc-token': 'PMxqO84bcPM5DgkX',
        'cloud_current_workspaceId': '63835346',
        '_wt': 'eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjc4NzYxMjc2fQ%3D%3D.32f4d13b0069c9f47f35e7e334a0d8404cab1411e1a1fc54363b998fa76f4ffb',
    }

    start_headers = {
        'Host': 'www.tapd.cn',
        # 'Cookie': '__root_domain_v=.tapd.cn; _qddaz=QD.271048197899749; 63835346bug_create_template=1163835346001000268; tcase_selected_category_id=1163835346001001771; new_worktable=search_filter; _t_uid=2039622907; _t_crop=22912571; tapd_div=101_3; tapdsession=1678441751d083b6e27a676b7bd3b8642229db7a7fdbf73939741830f8f3f02ef7f75c46f7; t_u=226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8ed6487938d984910ceb15bdf4a7b3bf01dc65e068b228a5fde9333e46ff239bffd38a2bfa9701f215295d673b8fabff94%7C1; t_cloud_login=myy%40addcn.com; dsc-token=PMxqO84bcPM5DgkX; cloud_current_workspaceId=63835346; _wt=eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjc4NzYxMjc2fQ%3D%3D.32f4d13b0069c9f47f35e7e334a0d8404cab1411e1a1fc54363b998fa76f4ffb',
        'sec-ch-ua': '"Google Chrome";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
        'sec-ch-ua-mobile': '?0',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
        'dsc-token': 'PMxqO84bcPM5DgkX',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'accept': '*/*',
        'x-requested-with': 'XMLHttpRequest',
        'sec-ch-ua-platform': '"Windows"',
        'origin': 'https://www.tapd.cn',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-mode': 'cors',
        'sec-fetch-dest': 'empty',
        'referer': 'https://www.tapd.cn/63835346/userviews/edit_show_fields/1000000000000000097?location=/bugtrace/buglists/query&height=641&width=800&KeepThis=true&TB_iframe=false&timsestamp=1678760976&rand=1678761052384',
        'accept-language': 'zh-CN,zh;q=0.9',
    }

    start_data = {
        'data[fields][id]': 'id',
        'data[fields][title]': 'title',
        'data[fields][status]': 'status',
        'data[fields][current_owner]': 'current_owner',
        'data[fields][reporter]': 'reporter',
        'data[fields][de]': 'de',
        'data[fields][te]': 'te',
        'data[fields][created]': 'created',
        'data[fields][in_progress_time]': 'in_progress_time',
        'data[fields][resolved]': 'resolved',
        'data[fields][closed]': 'closed',
        'data[fields][custom_field_one]': 'custom_field_one',
        'data[fields][custom_field_7]': 'custom_field_7',
        'data[fields][custom_field_8]': 'custom_field_8',
        'custom_fields': 'id;title;status;reporter;created;in_progress_time;resolved;closed;custom_field_one;de;te;custom_field_8;custom_field_7;current_owner;',
        'location': '/bugtrace/buglists/query',
        'workspace_id': '63835346',
        'workspace_code': '63835346',
        'id': '1000000000000000097',
    }

    start_response = requests.post('https://www.tapd.cn/63835346/userviews/edit_show_fields/', cookies=start_cookies, headers=start_headers, data=start_data)

    cookies = {
        '2039622907_63835346_/bugtrace/buglists/query_remember_view': '1163835346001033325',
        'bugtrace_view_63835346': '%5B%22project_id%22%2C%22title%22%2C%22current_owner%22%2C%22status%22%2C%22priority%22%2C%22severity%22%2C%22reporter%22%2C%22custom_field_three%22%2C%22custom_field_one%22%2C%22custom_field_7%22%2C%22created%22%2C%22bugtype%22%5D',
        '2039622907_63835346_/bugtrace/buglists/query_remember_view': '1163835346001033325',
        '__root_domain_v': '.tapd.cn',
        '_qddaz': 'QD.271048197899749',
        '63835346bug_create_template': '1163835346001000268',
        'tcase_selected_category_id': '1163835346001001771',
        'new_worktable': 'search_filter',
        '_t_uid': '2039622907',
        '_t_crop': '22912571',
        'tapd_div': '101_3',
        'tapdsession': '1678441751d083b6e27a676b7bd3b8642229db7a7fdbf73939741830f8f3f02ef7f75c46f7',
        't_u': '226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8ed6487938d984910ceb15bdf4a7b3bf01dc65e068b228a5fde9333e46ff239bffd38a2bfa9701f215295d673b8fabff94%7C1',
        't_cloud_login': 'myy%40addcn.com',
        'dsc-token': 'PMxqO84bcPM5DgkX',
        'cloud_current_workspaceId': '63835346',
        '_wt': 'eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjc4ODUyNTE2fQ%3D%3D.b64032fbaeda3b47bf2de007511f76f4ba51a1de1e6dcd5bd643a88e770678d7',
    }

    headers = {
        'Host': 'www.tapd.cn',
        'Cache-Control': 'max-age=0',
        'sec-ch-ua': '"Google Chrome";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'Upgrade-Insecure-Requests': '1',
        'Origin': 'https://www.tapd.cn',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Referer': 'https://www.tapd.cn/63835346/bugtrace/buglists/query/1/created/desc?query_token=2023031402cd90d3affdac288f2951214dc8bcf8',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        # 'Cookie': '2039622907_63835346_/bugtrace/buglists/query_remember_view=1163835346001033325; bugtrace_view_63835346=%5B%22project_id%22%2C%22title%22%2C%22current_owner%22%2C%22status%22%2C%22priority%22%2C%22severity%22%2C%22reporter%22%2C%22custom_field_three%22%2C%22custom_field_one%22%2C%22custom_field_7%22%2C%22created%22%2C%22bugtype%22%5D; 2039622907_63835346_/bugtrace/buglists/query_remember_view=1163835346001033325; __root_domain_v=.tapd.cn; _qddaz=QD.271048197899749; 63835346bug_create_template=1163835346001000268; tcase_selected_category_id=1163835346001001771; new_worktable=search_filter; _t_uid=2039622907; _t_crop=22912571; tapd_div=101_3; tapdsession=1678441751d083b6e27a676b7bd3b8642229db7a7fdbf73939741830f8f3f02ef7f75c46f7; t_u=226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8ed6487938d984910ceb15bdf4a7b3bf01dc65e068b228a5fde9333e46ff239bffd38a2bfa9701f215295d673b8fabff94%7C1; t_cloud_login=myy%40addcn.com; dsc-token=PMxqO84bcPM5DgkX; cloud_current_workspaceId=63835346; _wt=eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjc4ODUyNTE2fQ%3D%3D.b64032fbaeda3b47bf2de007511f76f4ba51a1de1e6dcd5bd643a88e770678d7',
    }

    params = {
        'query_token': '2023031402cd90d3affdac288f2951214dc8bcf8',
    }
    # 获取上个月的年月日期
    now = datetime.now()
    last_month = now - timedelta(days=now.day)
    year_month = last_month.strftime("%Y-%m")
    days_in_month = calendar.monthrange(last_month.year, last_month.month)[1]

    data = [
        ('workspace_name', 'T5的所有项目'),
        ('select_workspace', '63835346|35299590|51931447'),
        ('viewfilter[project_id][column]', 'project_id'),
        ('viewfilter[project_id][operator]', 'IN'),
        ('relationship[]', 'AND'),
        ('custom_fields', 'reporter;id;title;status;created;in_progress_time;resolved;closed;custom_field_one;de;te;custom_field_8;custom_field_7;current_owner;project_id'),
        ('viewfilter[2][column]', 'title'),
        ('viewfilter[2][operator]', 'LIKE'),
        ('viewfilter[2][value][]', ''),
        ('viewfilter[3][column]', 'current_owner'),
        ('viewfilter[3][operator]', 'LIKE'),
        ('viewfilter[3][value][]', ''),
        ('viewfilter[4][column]', 'status'),
        ('viewfilter[4][operator]', 'IN'),
        ('viewfilter[4][value][][]', 'empty'),
        ('viewfilter[4][value][][]', 'new'),
        ('viewfilter[4][value][][]', 'in_progress'),
        ('viewfilter[4][value][][]', 'resolved'),
        ('viewfilter[4][value][][]', 'verified'),
        ('viewfilter[4][value][][]', 'reopened'),
        ('viewfilter[4][value][][]', 'closed'),
        ('viewfilter[5][column]', 'priority'),
        ('viewfilter[5][operator]', 'IN'),
        ('viewfilter[6][column]', 'severity'),
        ('viewfilter[6][operator]', 'IN'),
        ('viewfilter[7][column]', 'reporter'),
        ('viewfilter[7][operator]', 'LIKE'),
        ('viewfilter[7][value][]', ''),
        ('viewfilter[8][column]', 'custom_field_three'),
        ('viewfilter[8][operator]', '='),
        ('viewfilter[9][column]', 'custom_field_one'),
        ('viewfilter[9][operator]', '='),
        ('viewfilter[10][column]', 'custom_field_7'),
        ('viewfilter[10][operator]', '='),
        ('viewfilter[10][value][][]', '日常BUG'),
        ('viewfilter[11][column]', 'created'),
        ('viewfilter[11][operator]', 'BETWEEN'),
        ('viewfilter[11][value][0][from]', f'{startTime} 00:00'),
        ('viewfilter[11][value][0][to]', f'{endTime} 23:59'),
        ('viewfilter[12][column]', 'bugtype'),
        ('viewfilter[12][operator]', 'IN'),
        ('data[submit]', '查询'),
        ('#####[column]', 'id'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'title'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'description'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'module'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'reporter'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'deadline'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'created'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'reopen_time'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'closed'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'modified'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'lastmodify'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'de'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'te'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'auditer'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'confirmer'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'current_owner'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'status'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'resolution'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'priority'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'severity'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'platform'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'os'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'testmode'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'testtype'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'testphase'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'source'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'frequency'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'cc'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'fixer'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'closer'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'participator'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'version_report'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'version_test'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'version_fix'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'version_close'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'baseline_find'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'baseline_join'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'baseline_close'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'baseline_test'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'originphase'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'sourcephase'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'bugtype'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'reject_time'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'in_progress_time'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'resolved'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'verify_time'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'assigned_time'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'iteration_id'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'project_id'),
        ('#####[value][]', ''),
        ('#####[column]', 'begin'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'due'),
        ('#####[operator]', 'BETWEEN'),
        ('#####[value][0][from]', ''),
        ('#####[value][0][to]', ''),
        ('#####[value][count]', '1'),
        ('#####[value][type]', 'month'),
        ('#####[value][]', ''),
        ('#####[column]', 'release_id'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'label'),
        ('#####[operator]', 'IN'),
        ('#####[column]', 'size'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'effort'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'effort_completed'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'remain'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'exceed'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'progress'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_one'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_two'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_three'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_four'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_five'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_6'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_7'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'custom_field_8'),
        ('#####[operator]', '='),
        ('#####[value][]', ''),
        ('#####[column]', 'BugStoryRelation_relative_id'),
        ('#####[operator]', 'LIKE'),
        ('#####[value][]', ''),
        ('#####[value][]', 'NULL'),
        ('#####[value][]', 'NULL'),
        ('#####[value][]', ''),
        ('dsc_token', 'PMxqO84bcPM5DgkX'),
    ]

    response = requests.post(
        'https://www.tapd.cn/63835346/bugtrace/buglists/query/1/created/desc',
        params=params,
        cookies=cookies,
        headers=headers,
        data=data,
    )

    excel_title = ["ID", "标题", "状态", "创建人", "创建时间", "接受处理时间", "解决时间", "关闭时间", "Bug等级", "开发人员", "测试人员", "所属业务线", "缺陷分类", "及时完成", "当天解决", "及时回归", "备注"]

    workbook = xlwt.Workbook('utf-8')
    sheet = workbook.add_sheet("Sheet1", cell_overwrite_ok=True)

    fileName = "-".join(fileName.split("-")[:2])
    fileName = fileName.split("-")[1]
    print("queryFileName---------", fileName)
    save_file = f"{fileName}日常Bug统计数据.xls"

    # 写入表头
    for column in range(len(excel_title)):
        sheet.write(0, column, excel_title[column])

    # bug编号
    bug_num_re = re.compile(r'<td id="id" class="editable-dropdown-label editable-click-area.*?<a href="(?P<bugHref>.*?)"><span>(?P<bugID>.*?)</span>', re.S)

    # bug标题
    title_re = re.compile(r"<span class='workitem-icon workitem-icon__bug' title='bug'></span>.*?<a href=.*?title=.*?>(?P<bugName>.*?)</a>", re.S)

    # bug状态
    status_re = re.compile(r'<td id="status" class="editable-dropdown-label editable-click-area.*?<a refresh_url=.*?>(?P<status_name>.*?)</a>', re.S)

    # bug创建人
    create_name_re = re.compile(r'<td id="reporter" class="editable-dropdown-label editable-click-area.*?<span class="editable-value" title=.*?>(?P<create_name>.*?)</span>', re.S)

    # bug创建时间
    create_time_re = re.compile(r'<td id="created" class="editable-dropdown-label editable-click-area.*?>(?P<create_time>.*?)</td>', re.S)

    # bug接受处理时间
    progress_time_re = re.compile(r'<td id="in_progress_time" class="editable-dropdown-label editable-click-area.*?>(?P<progress_time>.*?)</td>', re.S)

    # bug解决时间
    resolved_time_re = re.compile(r'<td id="resolved" class="editable-dropdown-label editable-click-area.*?>(?P<resolved_time>.*?)</td>', re.S)

    # bug关闭时间
    close_time_re = re.compile(r'<td id="closed" class="editable-dropdown-label editable-click-area.*?>(?P<close_time>.*?)</td>', re.S)

    # bug等级
    level_time_re = re.compile(r'<td id="custom_field_one" class="editable-dropdown-label editable-click-area.*?data-editable-field=.*?data-editable-value=(?P<level_time>.*?)>', re.S)

    # bug开发人员
    devpeople_re = re.compile(r'<td id="de" class="editable-dropdown-label editable-click-area.*?data-editable-field=.*?<span class="editable-value" title="(?P<devpeople>.*?)">', re.S)

    # bug测试人员
    testerpeople_re = re.compile(r'<td id="te" class="editable-dropdown-label editable-click-area.*?data-editable-field=.*?<span class="editable-value" title="(?P<testerpeople>.*?)">', re.S)

    # 业务线
    product_re = re.compile(r'data-editable-field="custom_field_8".*?<span.*?>(?P<product>.*?)</span>', re.S)

    # 处理人
    create_prople_re = re.compile(r'data-editable-field="current_owner".*?<span.*?>(?P<create_prople>.*?)</span>', re.S)

    bugNum_list = bug_num_re.finditer(response.text)
    bugNum_map = []
    global bugHref_map
    bugHref_map = []
    for x in bugNum_list:
        bugNum_map.append(x.group("bugID"))
        bugHref_map.append(str(x.group("bugHref")))

    print("bugHref_map ====", str(bugHref_map))
    # 写入ID
    for row in range(len(bugNum_map)):
        sheet.write(row + 1, 0, bugNum_map[row])

    fter = title_re.finditer(response.text)
    global title_map
    title_map = []
    for x in fter:
        title_map.append(x.group("bugName"))
        # print(x.group("bugName"))

    # 写入bug标题
    for row in range(len(title_map)):
        h = "https://www.tapd.cn" + bugHref_map[row]
        sheet.write(row + 1, 1, xlwt.Formula('HYPERLINK("{}"; "{}")'.format(h, title_map[row])))

    status_list = status_re.finditer(response.text)
    status_map = []
    for x in status_list:
        status_map.append(x.group("status_name"))
        print(x.group("status_name"))

    # 写入bug状态
    for row in range(len(status_map)):
        sheet.write(row + 1, 2, status_map[row])

    create_name_list = create_name_re.finditer(response.text)
    create_name_map = []
    for x in create_name_list:
        create_name_map.append(x.group("create_name").strip())
        print(x.group("create_name").strip())

    # 写入bug创建人
    for row in range(len(create_name_map)):
        sheet.write(row + 1, 3, create_name_map[row])

    create_time_list = create_time_re.finditer(response.text)
    create_time_map = []
    for x in create_time_list:
        create_time_map.append(x.group("create_time").strip())
        print(x.group("create_time").strip())

    # 写入bug创建时间
    for row in range(len(create_time_map)):
        sheet.write(row + 1, 4, create_time_map[row])

    progress_time_list = progress_time_re.finditer(response.text)
    progress_time_map = []
    for x in progress_time_list:
        progress_time_map.append(x.group("progress_time").strip())
        print(x.group("progress_time").strip())

    # 写入bug接受处理时间
    for row in range(len(progress_time_map)):
        sheet.write(row + 1, 5, progress_time_map[row])

    resolved_time_list = resolved_time_re.finditer(response.text)
    resolved_time_map = []
    for x in resolved_time_list:
        resolved_time_map.append(x.group("resolved_time").strip())
        print(x.group("resolved_time").strip())

    # 写入bug解决时间
    for row in range(len(resolved_time_map)):
        sheet.write(row + 1, 6, resolved_time_map[row])

    close_time_list = close_time_re.finditer(response.text)
    close_time_map = []
    for x in close_time_list:
        close_time_map.append(x.group("close_time").strip())
        print(x.group("close_time").strip())

    # 写入bug关闭时间
    for row in range(len(close_time_map)):
        sheet.write(row + 1, 7, close_time_map[row])

    level_time_list = level_time_re.finditer(response.text)
    level_time_map = []
    for x in level_time_list:
        level_time_map.append(x.group("level_time").replace("'", ""))
        print(x.group("level_time").replace("'", ""))

    # 写入bug等级
    for row in range(len(level_time_map)):
        sheet.write(row + 1, 8, level_time_map[row])

    devpeople_list = devpeople_re.finditer(response.text)
    devpeople_map = []
    for x in devpeople_list:
        devpeople_map.append(x.group("devpeople").replace("'", ""))
        print(x.group("devpeople").replace("'", ""))

    # 写入bug开发人员
    for row in range(len(devpeople_map)):
        sheet.write(row + 1, 9, devpeople_map[row])

    testerpeople_list = testerpeople_re.finditer(response.text)
    testerpeople_map = []
    for x in testerpeople_list:
        testerpeople_map.append(x.group("testerpeople").replace("'", ""))
        print(x.group("testerpeople").replace("'", ""))

    # 写入bug测试人员
    for row in range(len(testerpeople_map)):
        sheet.write(row + 1, 10, testerpeople_map[row])

    product_list = product_re.finditer(response.text)
    product_map = []
    for x in product_list:
        product_map.append(x.group("product").replace("'", ""))
        print(x.group("product").replace("'", ""))

    # 写入bug所属业务线
    for row in range(len(product_map)):
        sheet.write(row + 1, 11, product_map[row])

    create_prople_list = create_prople_re.finditer(response.text)
    create_prople_map = []
    for x in create_prople_list:
        create_prople_map.append(x.group("create_prople").replace("'", ""))
        print(x.group("create_prople").replace("'", ""))

    # 写入bug处理人
    for row in range(len(product_map)):
        sheet.write(row + 1, 12, "日常BUG")

    workbook.save(save_file)

    # thread.signal_progress_update.emit(30)


def readTapdList(startTime):
    # 获取上个月的年月日期
    now = datetime.now()
    last_month = now - timedelta(days=now.day)
    year_month = last_month.strftime("%Y-%m")
    days_in_month = calendar.monthrange(last_month.year, last_month.month)[1]

    print("start......", startTime)
    data = xlrd.open_workbook(f"{startTime}日常Bug统计数据.xls")
    sheet = data.sheets()[0]  # 获取第一sheet页

    newExcel = copy(data)

    newWs = newExcel.get_sheet(0)
    # 按索引号获取sheet的名字（string类型）

    row_number = sheet.nrows  # 获取所有行数量
    # print(row_number)

    row_data = []
    for cell in range(1, row_number):  # 从第二行开始读取
        value_list = sheet.row_values(cell)
        bug_create_time = value_list[4]
        bug_solve_time = value_list[6]
        bug_close_time = value_list[7]
        print(bug_create_time, bug_solve_time)
        level = value_list[8]

        if str(bug_solve_time).replace("\r\n\t", "") != "":
            start_time = datetime.strptime(bug_create_time, '%Y-%m-%d %H:%M')
            end_time = datetime.strptime(bug_solve_time, '%Y-%m-%d %H:%M')

            print(bug_create_time, bug_solve_time)
            time_diff = end_time - start_time
            # print("时间差为：", time_diff)
            # 定义日期格式
            date_format = "%Y-%m-%d %H:%M"

            # 判断时间范围内工作日的天数
            days_off = 5, 6
            workdays = [x for x in range(7) if x not in days_off]
            dayly = rrule.rrule(rrule.DAILY, dtstart=start_time, until=end_time, byweekday=workdays)
            days = dayly.count()
            print(days)

            if str(level).strip() == "1（当天）":
                if days <= 1:
                    newWs.write(cell, 13, "是")
                else:
                    newWs.write(cell, 13, "否")
            elif str(level).strip() == "2（3天内）":
                if days <= 3:
                    newWs.write(cell, 13, "是")
                else:
                    newWs.write(cell, 13, "否")
            elif str(level).strip() == "3（5天内）":
                if days <= 5:
                    newWs.write(cell, 13, "是")
                else:
                    newWs.write(cell, 13, "否")
            elif str(level).strip() == "4（7天内）":
                if days <= 7:
                    newWs.write(cell, 13, "是")
                else:
                    newWs.write(cell, 13, "否")
            elif str(level).strip() == "顽固（180 天）":
                newWs.write(cell, 13, "是")

            if start_time.date() == end_time.date():
                newWs.write(cell, 14, "是")
            else:
                newWs.write(cell, 14, "否")

            if str(bug_close_time).replace("\r\n\t", "") != "":
                close_time = datetime.strptime(bug_close_time, '%Y-%m-%d %H:%M')
                close_solve_time = rrule.rrule(rrule.DAILY, dtstart=end_time, until=close_time, byweekday=workdays)
                daydiff = close_solve_time.count()
                if daydiff <= 2:
                    newWs.write(cell, 15, "是")
                else:
                    newWs.write(cell, 15, "否")
            else:
                newWs.write(cell, 15, "否")

        else:
            newWs.write(cell, 13, "否")
            newWs.write(cell, 14, "否")
            newWs.write(cell, 15, "是")

    newExcel.save(f"{startTime}日常Bug统计数据.xls")


def writeCount(startTime):
    # 获取上个月的年月日期
    now = datetime.now()
    last_month = now - timedelta(days=now.day)
    year_month = last_month.strftime("%Y-%m")
    days_in_month = calendar.monthrange(last_month.year, last_month.month)[1]

    print("write...", startTime)
    data = xlrd.open_workbook(f"{startTime}日常Bug统计数据.xls")

    sheet = data.sheets()[0]

    row_number = sheet.nrows

    newExcel = copy(data)

    newWs = newExcel.get_sheet(0)

    # 获取第一个工作表
    worksheet = data.sheet_by_index(0)

    # 获取列名所在行的索引
    header_row_index = 0

    # 获取列名
    header_row = worksheet.row(header_row_index)
    header = [cell.value for cell in header_row]

    # 获取 ColumnA 列的索引
    column_index = header.index('及时完成')
    time_back_column_index = header.index('及时回归')
    today_column_index = header.index('当天解决')

    # 筛选符合条件（及时解决：是）的数据
    filtered_rows = [row_index for row_index in range(1, worksheet.nrows) if worksheet.cell_value(row_index, column_index) == "是"]

    # 筛选符合条件（及时解决：否）的数据
    filtered_time_out_rows = [row_index for row_index in range(1, worksheet.nrows) if worksheet.cell_value(row_index, column_index) == "否"]

    # 筛选符合条件（及时回归：是）的数据
    filtered_time_back_rows = [row_index for row_index in range(1, worksheet.nrows) if worksheet.cell_value(row_index, time_back_column_index) == "是"]

    # 筛选符合条件（及时回归：否）的数据
    filtered_time_out_back_rows = [row_index for row_index in range(1, worksheet.nrows) if worksheet.cell_value(row_index, time_back_column_index) == "否"]

    # 筛选符合条件（当天解决：是）的数据
    filtered_today_rows = [row_index for row_index in range(1, worksheet.nrows) if worksheet.cell_value(row_index, today_column_index) == "是"]

    time_out_value_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')

    for time_out_solve_row_index in filtered_time_out_rows:
        newWs.write(int(time_out_solve_row_index), 13, "否", style=time_out_value_style)

    # for filtered_today_time_out_index in filtered_today_time_out_rows:
    #     newWs.write(int(filtered_today_time_out_index), 14, "否", style=time_out_value_style)

    for time_out_back_row_index in filtered_time_out_back_rows:
        newWs.write(int(time_out_back_row_index), 15, "否", style=time_out_value_style)

    style = xlwt.easyxf('font: underline single, color blue;')
    # 写入标题超链接
    for index in range(1, worksheet.nrows):
        h = "https://www.tapd.cn" + str(bugHref_map[index - 1])
        hyperlink_formula = xlwt.Formula('HYPERLINK("{}"; "{}")'.format(h, title_map[index - 1]))
        newWs.write(index, 1, hyperlink_formula, style=style)

    # 统计数量  及时解决数
    count = len(filtered_rows)

    # 统计数量  超时解决数
    time_out_count = len(filtered_time_out_rows)

    # 统计数量  及时回归数
    time_back_count = len(filtered_time_back_rows)

    # 统计数量  超时回归数
    time_back_out_count = len(filtered_time_out_back_rows)

    # 统计数量   当天解决数
    time_today_count = len(filtered_today_rows)

    # 统计数量  及时解决率
    TimelyResolutionRate = '{:.2%}'.format(count / (worksheet.nrows - 1))

    # 统计数量  超时解决率
    TimeOutResolutionRate = '{:.2%}'.format(time_out_count / (worksheet.nrows - 1))

    # 统计数量  及时回归率
    TimeBackRate = '{:.2%}'.format(time_back_count / (worksheet.nrows - 1))

    # 统计数量  超时回归率
    TimeOutBackRate = '{:.2%}'.format(time_back_out_count / (worksheet.nrows - 1))

    # 统计数量  当天解决率
    TimeTodayRate = '{:.2%}'.format(time_today_count / (worksheet.nrows - 1))

    print(
        f"及时解决数为:{count},及时解决率为：{TimelyResolutionRate},超时解决数为:{time_out_count},超时解决率为：{TimeOutResolutionRate},及时回归数为:{time_back_count},及时回归率为：{TimeBackRate},超时回归数为：{time_back_out_count}，超时回归率为:{TimeOutBackRate},当天解决数：{time_today_count},当天解决率：{TimeTodayRate}")

    borders = xlwt.Borders()

    borders.left = xlwt.Borders.MEDIUM
    borders.right = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.top = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.bottom = xlwt.Borders.MEDIUM

    style = XFStyle()
    pattern = Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # pattern.pattern_fore_colour = Style.colour_map['#fce5cd']
    pattern.pattern_fore_colour = 47
    style.pattern = pattern
    style.borders = borders

    value_style = XFStyle()
    value_style.borders = borders

    newWs.write(row_number + 3, 3, "BUG总数", style=style)
    newWs.write(row_number + 3, 4, row_number - 1, style=value_style)
    newWs.write(row_number + 3, 5, "按时解决总数", style=style)
    newWs.write(row_number + 3, 6, count, style=value_style)

    newWs.write(row_number + 4, 3, "及时解决数", style=style)
    newWs.write(row_number + 4, 4, count, style=value_style)
    newWs.write(row_number + 4, 5, "按时解决率", style=style)
    newWs.write(row_number + 4, 6, TimelyResolutionRate, style=value_style)

    newWs.write(row_number + 5, 3, "超时解决数", style=style)
    newWs.write(row_number + 5, 4, time_out_count, style=value_style)
    newWs.write(row_number + 5, 5, "超时解决率", style=style)
    newWs.write(row_number + 5, 6, TimeOutResolutionRate, style=value_style)

    newWs.write(row_number + 6, 3, "当天解决数", style=style)
    newWs.write(row_number + 6, 4, time_today_count, style=value_style)
    newWs.write(row_number + 6, 5, "当天解决率", style=style)
    newWs.write(row_number + 6, 6, TimeTodayRate, style=value_style)

    newWs.write(row_number + 7, 3, "及时回归", style=style)
    newWs.write(row_number + 7, 4, time_back_count, style=value_style)
    newWs.write(row_number + 7, 5, "及时回归率", style=style)
    newWs.write(row_number + 7, 6, TimeBackRate, style=value_style)

    newWs.write(row_number + 8, 3, "超时回归", style=style)
    newWs.write(row_number + 8, 4, time_back_out_count, style=value_style)
    newWs.write(row_number + 8, 5, "超时回归率", style=style)
    newWs.write(row_number + 8, 6, TimeOutBackRate, style=value_style)

    newExcel.save(f"{startTime}日常Bug统计数据.xls")

    # thread.signal_progress_update.emit(80)


def createPlanRequest(obj, yewustr, title, author, project, xuqiu_id, startTime, endTime):
    import requests
    cookies = {
        '2039622907_63835346_/sparrow/test_plan/plan_list_remember_view': '1163835346001032377',
        'plan_list1163835346001032377sort': 'created%7Edesc',
        'tui_filter_fields': '%5B%22name%22%2C%22version%22%2C%22owner%22%2C%22status%22%5D',
        'tflnewfilter_status': 'open',
        'recentUserCookie2': 'true',
        '_qddaz': 'QD.271048197899749',
        '63835346bug_create_template': '1163835346001000268',
        'tcase_selected_category_id': '1163835346001001771',
        'tapdsession': '1683163754c73cd18a5b15c71b40ef513c56c661868c48836adaa4b863cebf969b706a1eb7',
        '__root_domain_v': '.tapd.cn',
        't_u': '226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8e8453a661d3c4e697c0ea4204ddcce7ad1c6ed51c49cbd376124c1616836cd948bbaa7675644b772faa63425405d8bde6%7C1',
        'dsc-token': 'XJaRQLxYWyYmQ6UF',
        '_t_uid': '2039622907',
        '_t_crop': '22912571',
        'tapd_div': '101_3',
        'cloud_current_workspaceId': '63835346',
        'new_worktable': 'search_filter',
        'recentUserCookie': 'true',
        '_wt': 'eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjkwNTM3NTMyfQ%3D%3D.67c64d5cf6947412560e47fb3ecdc7e1cbd4376eb04afb22ca818cad9457aa57',
    }

    headers = {
        'Host': 'www.tapd.cn',
        'Cache-Control': 'max-age=0',
        'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'Upgrade-Insecure-Requests': '1',
        'Origin': 'https://www.tapd.cn',
        'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundaryCURkIMVWjE2kgO3M',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Referer': 'https://www.tapd.cn/63835346/sparrow/test_plan/add?action_timestamp=16762924',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        # 'Cookie': '2039622907_63835346_/sparrow/test_plan/plan_list_remember_view=1163835346001032377; plan_list1163835346001032377sort=created%7Edesc; tui_filter_fields=%5B%22name%22%2C%22version%22%2C%22owner%22%2C%22status%22%5D; tflnewfilter_status=close; recentUserCookie2=true; _qddaz=QD.271048197899749; 63835346bug_create_template=1163835346001000268; tcase_selected_category_id=1163835346001001771; tapdsession=1683163754c73cd18a5b15c71b40ef513c56c661868c48836adaa4b863cebf969b706a1eb7; __root_domain_v=.tapd.cn; t_u=226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8e8453a661d3c4e697c0ea4204ddcce7ad1c6ed51c49cbd376124c1616836cd948bbaa7675644b772faa63425405d8bde6%7C1; dsc-token=cX5pESUv4OyfJABh; _t_uid=2039622907; _t_crop=22912571; tapd_div=101_3; cloud_current_workspaceId=63835346; new_worktable=search_filter; recentUserCookie=true; _wt=eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjkwNTM3NTMyfQ%3D%3D.67c64d5cf6947412560e47fb3ecdc7e1cbd4376eb04afb22ca818cad9457aa57',
    }

    params = {
        'action_timestamp': '53531587',
    }

    data = f'------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="add_test_plan_token"\n\n2039622907174115597\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][template_id]"\n\n1163835346001000190\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][id]"\n\n\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][name]"\n\n{title}\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][description]"\n\n<div class="lake-content" typography="classic"><p class="ne-p"><span style="font-size:large;"><strong><span class="ne-text">测试目标</span></strong><span class="ne-text"> </span></span></p><p class="ne-p"><strong><span class="ne-text">{obj["test_target"]}</span></strong></p><p class="ne-p"><strong><span class="ne-text"><br/></span></strong></p><p class="ne-p"><strong><span class="ne-text"><br/></span></strong></p><p class="ne-p"><span class="ne-text"></span></p><p class="ne-p"><span class="ne-text"></span></p><p class="ne-p"><strong><span class="ne-text"><span style="font-size:large;">测试范围</span></span></strong></p><table class="ne-table"><tbody><tr><td width="150"><p class="ne-p"><span class="ne-text">业务名称</span></p></td><td width="150"><p class="ne-p"><span class="ne-text"> 页面/功能</span></p></td><td width="150"><p class="ne-p"><span class="ne-text"> 模块</span></p></td><td width="150"><p class="ne-p"><span class="ne-text"> 端口</span></p></td><td width="150"><p class="ne-p"><span class="ne-text"> 负责人</span></p></td></tr>{yewustr}</tbody></table><p class="ne-p"><strong><span class="ne-text"><span style="font-size:large;">测试进度</span></span></strong></p><table class="ne-table"><tbody><tr><td width="250"><p class="ne-p"><span class="ne-text">阶段 </span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 预估时间</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 实际时间</span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">用例设计时长</span><span class="ne-text" style="color: rgb(140, 140, 140);">（项目设计用例耗时）</span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["case_time"]}</p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["case_shiji_time"]}</span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">用例评审日期</span><span class="ne-text" style="color: rgb(140, 140, 140);">（项目用例评审会召开时间）</span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["pingshen_time"]}</p></td><td width="250"><p class="ne-p">&nbsp;{obj["pingshen_shiji_time"]}</p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">提测日期</span><span class="ne-text" style="color: rgb(140, 140, 140);">（项目对应端口提测时间）</span></p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["tice_time"]}</span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["tice_shiji_time"]}</p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">第一轮测试时长</span><span class="ne-text" style="color: rgb(140, 140, 140);">（执行第一轮测试用例总耗时）</span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["yilun_time"]}</p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["yilun_shiji_time"]}</span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">产品设计验收时间</span><span class="ne-text" style="color: rgb(140, 140, 140);">（测试完成后提交给产品及设计验收时间）</span><span class="ne-text"></span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["yanshou_time"]}</p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["yanshou_shiji_time"]}</span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text">上线时间 </span><span class="ne-text" style="color: rgb(140, 140, 140);">（项目具体上线时间）</span></p></td><td width="250"><p class="ne-p">&nbsp;{obj["shangxian_time"]}</p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["shangxian_shiji_time"]}</span></p></td></tr></tbody></table><p class="ne-p"><strong><span class="ne-text"><span style="font-size:large;">参考文档</span></span></strong></p><table class="ne-table"><tbody><tr><td width="375"><p class="ne-p"><span class="ne-text"> 类型 </span></p></td><td width="375"><p class="ne-p"><span class="ne-text"> 说明</span></p></td></tr><tr><td width="375"><p class="ne-p"><span class="ne-text">需求说明文档</span></p></td><td width="375"><p class="ne-p"><span class="ne-text">&nbsp;{obj["xuqiu_link"]}</span></p></td></tr><tr><td width="375"><p class="ne-p"><span class="ne-text">设计稿</span></p></td><td width="375"><p class="ne-p"><span class="ne-text">&nbsp;{obj["sheji_link"]}</span></p></td></tr><tr><td width="375"><p class="ne-p"><span class="ne-text">API文档</span></p></td><td width="375"><p class="ne-p"><span class="ne-text">&nbsp;{obj["api_link"]}</span></p></td></tr><tr><td width="375"><p class="ne-p"><span class="ne-text">测试环境</span></p></td><td width="375"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr><td width="375"><p class="ne-p"><span class="ne-text">其他</span></p></td><td width="375"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr></tbody></table><p class="ne-p"><strong><span class="ne-text"><span style="font-size:large;">测试策略</span></span></strong></p><table class="ne-table"><tbody><tr style="height: 49px;"><td width="250"><p class="ne-p"><span class="ne-text"> 类型</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 执行结果<br/></span><span class="ne-text"> （是/否）</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 备注</span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 功能测试</span></p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["gongneng_result"]}</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 接口测试<br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["api_result"]}</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 用户界面测试<br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text">&nbsp;{obj["jiemian_result"]}</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 性能测试</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250" style="height: auto;"><p class="ne-p"><span class="ne-text"> 安装测试<br/></span></p></td><td width="250" style="height: auto;"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250" style="height: auto;"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 弱网测试</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr style="height: 39px;"><td width="250"><p class="ne-p"><span class="ne-text"> 兼容性测试</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr></tbody></table><p class="ne-p"><strong><span class="ne-text"><span style="font-size:large;">回归范围</span></span></strong></p><table class="ne-table"><tbody><tr><td width="250"><p class="ne-p"><span class="ne-text"> 页面/功能</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 回归结果<br/></span><span class="ne-text"> （是/否）</span></p></td><td width="250"><p class="ne-p"><span class="ne-text"> 负责人</span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr><tr><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td><td width="250"><p class="ne-p"><span class="ne-text"><br/></span></p></td></tr></tbody></table></div>\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][iteration_id]"\n\n{project}\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][version]"\n\n版本1\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][status]"\n\nopen\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][type]"\n\n功能测试\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][start_date]"\n\n{startTime}\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][end_date]"\n\n{endTime}\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[TestPlan][owner]"\n\nT5{author};\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="data[Attachment][file1]"; filename=""\nContent-Type: application/octet-stream\n\n\n------WebKitFormBoundaryCURkIMVWjE2kgO3M\nContent-Disposition: form-data; name="dsc_token"\n\ncX5pESUv4OyfJABh\n------WebKitFormBoundaryCURkIMVWjE2kgO3M--\n'.encode()

    response = requests.post(
        'https://www.tapd.cn/63835346/sparrow/test_plan/save_and_return/security',
        params=params,
        cookies=cookies,
        headers=headers,
        data=data,
    )

    plan_id = response.url.split("=")[1]
    guanlianxuqiu(plan_id, xuqiu_id)


def guanlianxuqiu(plan_id, xuqiu_id):
    import requests
    cookies = {
        '2039622907_63835346_/sparrow/test_plan/plan_list_remember_view': '1163835346001032377',
        'plan_list1163835346001032377sort': 'created%7Edesc',
        'tui_filter_fields': '%5B%22name%22%2C%22version%22%2C%22owner%22%2C%22status%22%5D',
        'tflnewfilter_status': 'close',
        'tcase_list_63835346_filter_fields': '%5B%22status%22%2C%22name%22%5D',
        '_qddaz': 'QD.271048197899749',
        '63835346bug_create_template': '1163835346001000268',
        'tcase_selected_category_id': '1163835346001001771',
        'tapdsession': '1683163754c73cd18a5b15c71b40ef513c56c661868c48836adaa4b863cebf969b706a1eb7',
        '__root_domain_v': '.tapd.cn',
        't_u': '226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8e8453a661d3c4e697c0ea4204ddcce7ad1c6ed51c49cbd376124c1616836cd948bbaa7675644b772faa63425405d8bde6%7C1',
        'dsc-token': 'cX5pESUv4OyfJABh',
        'cloud_current_workspaceId': '63835346',
        'new_worktable': 'search_filter',
        'recentUserCookie': 'true',
        '_t_uid': '2039622907',
        '_t_crop': '22912571',
        'tapd_div': '101_3',
        '_wt': 'eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjkxMDI3Mzk4fQ%3D%3D.7857579e05eb1c9e05c12c1e082b2f22a82655a70f2a0a7e7490483aa42b66e8',
    }

    headers = {
        'Host': 'www.tapd.cn',
        'sec-ch-ua': '"Not/A)Brand";v="99", "Google Chrome";v="115", "Chromium";v="115"',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36',
        'DSC-TOKEN': 'cX5pESUv4OyfJABh',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Accept': '*/*',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua-platform': '"Windows"',
        'Origin': 'https://www.tapd.cn',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://www.tapd.cn/63835346/sparrow/test_plan/detail/1163835346001001904?action_timestamp=98373925',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        # 'Cookie': '2039622907_63835346_/sparrow/test_plan/plan_list_remember_view=1163835346001032377; plan_list1163835346001032377sort=created%7Edesc; tui_filter_fields=%5B%22name%22%2C%22version%22%2C%22owner%22%2C%22status%22%5D; tflnewfilter_status=close; tcase_list_63835346_filter_fields=%5B%22status%22%2C%22name%22%5D; _qddaz=QD.271048197899749; 63835346bug_create_template=1163835346001000268; tcase_selected_category_id=1163835346001001771; tapdsession=1683163754c73cd18a5b15c71b40ef513c56c661868c48836adaa4b863cebf969b706a1eb7; __root_domain_v=.tapd.cn; t_u=226ff5b7e565c664fb0500757d2cd9622aa23f9d813bdd8e8453a661d3c4e697c0ea4204ddcce7ad1c6ed51c49cbd376124c1616836cd948bbaa7675644b772faa63425405d8bde6%7C1; dsc-token=cX5pESUv4OyfJABh; cloud_current_workspaceId=63835346; new_worktable=search_filter; recentUserCookie=true; _t_uid=2039622907; _t_crop=22912571; tapd_div=101_3; _wt=eyJ1aWQiOiIyMDM5NjIyOTA3IiwiY29tcGFueV9pZCI6IjIyOTEyNTcxIiwiZXhwIjoxNjkxMDI3Mzk4fQ%3D%3D.7857579e05eb1c9e05c12c1e082b2f22a82655a70f2a0a7e7490483aa42b66e8',
    }

    data = {
        'data[story_ids][]': [
            xuqiu_id,
        ],
    }

    response = requests.post(
        f'https://www.tapd.cn/63835346/sparrow/test_plan/save_plan_story/{plan_id}',
        cookies=cookies,
        headers=headers,
        data=data,
    )

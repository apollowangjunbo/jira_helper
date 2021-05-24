from jira import JIRA
import os, time, xlrd

import logging
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)
formatter = logging.Formatter("%(asctime)s - %(levelname)s: %(message)s")
# 文件日志
log_path = os.path.join(os.path.dirname(__file__), 'log.txt')
if not os.path.exists(log_path):
	with open(log_path, 'w'): pass
fh = logging.FileHandler(log_path, mode='w')
fh.setLevel(logging.DEBUG)
fh.setFormatter(formatter)
logger.addHandler(fh)
# 控制台日志
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)  # 输出到console的log等级的开关
ch.setFormatter(formatter)
logger.addHandler(ch)

class Task:
    duedate = ''
    summary = ''
    description = ''
    epic = ''
    story_point = -1
    issue_type = '任务'
    time_tracking = ''
    story_type = '无'
    id = ''

    def __repr__(self):
        return f"duedate:{self.duedate},\nsummary:{self.summary},\ndescription:{self.description},\n" \
               f"epic:{self.epic},\nstory_point:{self.story_point},\ntime_tracking:{self.time_tracking},\n" \
               f"issue_type:{self.issue_type},\id:{self.id}\nstory_type:{self.story_type}\n"

    def __str__(self):
        return self.__repr__()


class SubTask:
    parent = -1
    duedate = ''
    summary = ''
    description = ''
    time_tracking = ''
    components = ''

    def __repr__(self):
        return f"parent:{self.parent},\nduedate:{self.duedate},\nsummary:{self.summary},\n" \
               f"description:{self.description},\ncomponents:{self.components}\ntime_tracking:{self.time_tracking},\n"

    def __str__(self):
        return self.__repr__()


class ExcelReader:
    user = ''
    password = ''
    sprintID = -1
    version = ''
    feature_team = ''
    tasks = dict()
    sub_tasks = []

    def __init__(self, path):
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        nrows = sh.nrows
        row_index = 0

        assert '账号' in str(sh.cell_value(row_index, 0)) and str(sh.cell_value(row_index, 1)).strip() != ''
        self.user = str(sh.cell_value(row_index, 1)).strip()
        row_index += 1

        assert '密码' in str(sh.cell_value(row_index, 0)) and str(sh.cell_value(row_index, 1)).strip() != ''
        self.password = str(sh.cell_value(row_index, 1)).strip()
        row_index += 1

        assert 'sprintID' in str(sh.cell_value(row_index, 0))
        if str(sh.cell_value(row_index, 1)).strip() != '':
            self.sprintID = int(sh.cell_value(row_index, 1))
        row_index += 1

        assert 'version' in str(sh.cell_value(row_index, 0))
        self.version = str(sh.cell_value(row_index, 1)).strip()
        row_index += 1

        assert 'FeatureTeam' in str(sh.cell_value(row_index, 0))
        self.feature_team = str(sh.cell_value(row_index, 1)).strip()
        row_index += 1

        # 空行
        while (row_index < nrows
               and '任务' not in str(sh.cell_value(row_index, 0))
               and '用户故事' not in str(sh.cell_value(row_index, 0))):
            row_index += 1
        row_index += 2
        if row_index >= nrows:
            print('没找到要创建的任务！')
        # 任务
        while sh.cell_value(row_index, 0) != '' and '子任务' not in str(sh.cell_value(row_index, 0)):
            self.load_task(sh.row(row_index))
            row_index += 1
        # 空行
        while row_index < nrows and '子任务' not in str(sh.cell_value(row_index, 0)):
            row_index += 1
        row_index += 2
        if row_index >= nrows:
            print('没找到要创建的子任务')
        # 子任务
        while row_index < nrows and sh.cell_value(row_index, 0) != '':
            self.load_subtask(sh.row(row_index))
            row_index += 1

    def get_date(self, cell: xlrd.sheet.Cell):
        if cell.ctype == 3:
            return xlrd.xldate_as_datetime(cell.value, 0).strftime("%Y-%m-%d")
        if '/' in cell.value:
            return time.strftime("%Y-%m-%d", time.strptime(cell.value.strip(), "%Y/%m/%d"))
        return cell.value.strip()

    def load_task(self, row):
        t = Task()
        t.duedate = self.get_date(row[1])
        t.summary = str(row[2].value).strip().replace('\n','')
        if t.summary == '':
            return
        t.description = str(row[3].value).strip()
        t.epic = str(row[4].value).strip()
        t.story_point = float(row[5].value)
        t.issue_type = '用户故事' if ('用户故事' in str(row[6].value).strip()) else '任务'
        t.time_tracking = str(row[7].value).strip()
        t.story_type = str(row[8].value).strip()
        self.tasks[int(row[0].value)] = t

    def load_subtask(self, row):
        t = SubTask()
        t.parent = int(row[0].value)
        t.duedate = self.get_date(row[1])
        t.summary = str(row[2].value).strip().replace('\n','')
        if t.summary == '':
            return
        t.description = str(row[3].value).strip()
        t.time_tracking = str(row[4].value).strip()
        t.components = str(row[5].value).strip()
        self.sub_tasks.append(t)


class JiraHelper:
    data = None
    jira = None

    def __init__(self, data: ExcelReader):
        self.data = data
        self.jira = JIRA(server='http://pm.glodon.com/newjira/', basic_auth=(data.user, data.password))

    def create_issues(self):
        for taskid in self.data.tasks:
            task = self.data.tasks[taskid]
            logger.debug(task)
            issue_dict = {
                'project': {'key': 'GMP'},  # G项目
                'issuetype': {'id': '10002' if task.issue_type == '任务' else '10001'},
                'summary': task.summary,  # 题目
                'description': task.description,
                'assignee': {'name': self.data.user},
            }
            if self.data.sprintID>0:
                issue_dict['customfield_10001'] = self.data.sprintID
            if self.data.feature_team != '':
                issue_dict['customfield_10007'] = {'value': self.data.feature_team}
            if self.data.version != '':
                issue_dict['fixVersions'] = [{'name': self.data.version}]
                issue_dict['versions'] = [{'name': self.data.version}]

            if task.duedate != '':
                issue_dict['duedate'] = task.duedate
            if task.story_point > 0:
                issue_dict['customfield_10006'] = task.story_point
            if task.epic != '':
                issue_dict['customfield_10002'] = task.epic
            if task.issue_type == '任务' and task.time_tracking != '':
                issue_dict['timetracking'] = {'originalEstimate': task.time_tracking, 'remainingEstimate': task.time_tracking}
            if task.issue_type =='用户故事':
                issue_dict['customfield_10214'] = {'value':task.story_type}
            logger.debug(issue_dict)
            new_issue = self.jira.create_issue(fields=issue_dict)
            self.data.tasks[taskid].id=new_issue.id
            logger.info(f"创建{task.issue_type}成功，名称：{task.summary}\t，截止日期：{task.duedate}\t，ID：{new_issue.key}")

        for subtask in self.data.sub_tasks:
            logger.debug(subtask)
            assert subtask.parent in self.data.tasks
            issue_dict = {
                'project': {'key': 'GMP'},  # G项目
                'issuetype': {'id': '10003'},
                'summary': subtask.summary,  # 题目
                'description': subtask.description,
                'assignee': {'name': self.data.user},
                'components':[{'name':subtask.components}],
                'parent': {'id': str(self.data.tasks[subtask.parent].id)},
            }

            if self.data.feature_team != '':
                issue_dict['customfield_10007'] = {'value': self.data.feature_team}
            if self.data.version != '':
                issue_dict['fixVersions'] = [{'name': self.data.version}]
                issue_dict['versions'] = [{'name': self.data.version}]

            if task.duedate != '':
                issue_dict['duedate'] = subtask.duedate
            if subtask.time_tracking != '':
                issue_dict['timetracking'] = {'originalEstimate': subtask.time_tracking,
                                              'remainingEstimate': subtask.time_tracking}
            logger.debug(issue_dict)
            sub_task = self.jira.create_issue(fields=issue_dict)
            logger.info(f"创建子任务成功，名称：{subtask.summary}，\t截止日期：{subtask.duedate}，\tID：{sub_task.key}")

    def jql_webpage(self):
        jql = f'assignee={self.data.user} And project="GMP"'
        if self.data.sprintID != '':
            jql += f' And Sprint={self.data.sprintID}'
        if self.data.feature_team!= '':
            jql += f' And "Feature Team"={self.data.feature_team}'
        if self.data.version!= '':
            jql += f' And fixVersion={self.data.version} And affectedVersion={self.data.version}'
        logger.info(f"创建结果查询页面： http://pm.glodon.com/newjira/issues/?jql={jql}")
        logger.info(f"成功！")


if __name__ == "__main__":
    data = ExcelReader(os.path.join(os.getcwd(), 'input.xls'))

    # print(f"user:{data.user},\npassword:{data.password},\nsprintID:{data.sprintID},\n"
    #       f"version:{data.version},\nfeature_team{data.feature_team}\n")
    # print(data.tasks)
    # print(data.sub_tasks)

    j = JiraHelper(data)
    j.create_issues()
    j.jql_webpage()
    input()

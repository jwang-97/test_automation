import re
import os
import datetime

def search(path, s):
    pathlist = []
    pathlist = []
    for x in os.listdir(path):
        fp = os.path.join(path, x)
        if os.path.isfile(fp) and s in x:
            pathlist.append(fp)
    return pathlist
class libsummary(object):
    def __init__(self,libcore,libabnormal,libnofile,libdiff):
        self.libcore=libcore
        self.libabnormal =libabnormal
        self.libnofile=libnofile
        self.libdiff = libdiff

class nosummarycase(object):
    def __init__(self,nosummaryfiles,nosummarycase):
        self.nosummaryfiles=nosummaryfiles
        self.nosummarycase=nosummarycase

class diffcase(object):
    def __init__(self,difffile,diffcases):
        self.difffile=difffile
        self.diffcases=diffcases

class subtask(object):
    def __init__(self,subtaskname,subtaskstatus,subtaskresult):
        self.subtaskname=subtaskname
        self.subtaskstatus=subtaskstatus
        self.subtaskresult=subtaskresult

class newtask(object):
    def __init__(task, track_number, benchmark, subtask):
        task.track_number=track_number
        task.benchmark=benchmark
        task.subtask=subtask

class case_cls(object):
    def __init__(self, cpu1_number, cpu1_changetime, cpu1_restart, cpu1_cost, cpu4_number,
                 cpu4_changetime, cpu4_restart, cpu4_cost, cpu4check1_number, cpu4check1_changetime,
                 cpu4check1_restart, cpu4check1_cost,total_case,total_changetime,total_restart,total_cost):
        self.cpu1_number = cpu1_number
        self.cpu1_changetime = cpu1_changetime
        self.cpu1_restart = cpu1_restart
        self.cpu1_cost = cpu1_cost
        self.cpu4_number = cpu4_number
        self.cpu4_changetime = cpu4_changetime
        self.cpu4_restart = cpu4_restart
        self.cpu4_cost = cpu4_cost
        self.cpu4check1_number = cpu4check1_number
        self.cpu4check1_changetime = cpu4check1_changetime
        self.cpu4check1_restart = cpu4check1_restart
        self.cpu4check1_cost = cpu4check1_cost
        self.total_case =total_case
        self.total_changetime=total_changetime
        self.total_restart=total_restart
        self.total_cost=total_cost

class core(object):
    def __init__(self,corecase,corefunction,coreexe,coreinfo):
        self.corecase = corecase
        self.corefunction=corefunction
        self.coreexe=coreexe
        self.coreinfo=coreinfo

class errorstr(object):
    def __init__(self,errorfunctionname,totalerrorname,errorcase):
        self.errorfunctionname=errorfunctionname
        self.totalerrorname=totalerrorname
        self.errorcase=errorcase


class DateParser(object):
    def __init__(self):
        self.pattern = re.compile(
            r'^((?:19|20)?\d{2})[-.]?((?:[0-1]?|1)[0-9])[-.]?((?:[0-3]?|[1-3])[0-9])?$'
        )

    def parse(self, strdate):
        m = self.pattern.match(strdate)
        flags = [False, False, False]
        if m:
            matches = list(m.groups())
            flags = list(map(lambda x: True if x != None else False, matches))
            results = list(map(lambda x: int(x) if x != None else 1, matches))
            # results = list(map(lambda x:1 if x==None else x, results))
            if results[0] < 100:
                if results[0] > 9:
                    results[0] += 1900
                else:
                    results[0] += 2000

            return (datetime.date(results[0], results[1], results[2]), flags)
        else:
            return (None,flags)


class dataobject(object):
    def __init__(self, diff, nofile):
        self.diff = diff
        self.nofile = nofile

class readautocase(object):
    def __init__(self,autocasereusltpath):
        self.autocaseresultpath=autocasereusltpath

    def parsefiles(self, filename):
        autocasediff = filename
        f = open(autocasediff)
        lines = f.readlines()
        finaldiff = []

        difflist = []
        onediff = []
        nofile = []

        for line in lines:
            if "diff -r" in line:
                if len(onediff) > 0:
                    difflist.append(onediff)
                    onediff = []
            if "Only in" in line:
                nofile.append(line)
                continue
            onediff.append(line)
        if len(onediff) > 0:
            difflist.append(onediff)

        for diff in difflist:
            current = ''
            benchmark = ''
            currentdate=''
            benchmarkdate=''
            for diffline in diff:
                if '<' in diffline:
                    current = current + diffline
                    currentsplits = current.split('<')
                    currentdate = currentsplits[1].strip()
                    if len(currentdate) > 0 and '"Value":' in currentdate:
                        currentdatesplit = currentdate.split('"Value":')
                        currentdate = currentdatesplit[1].strip()
                    if len(currentdate) > 0 and '"Version":' in currentdate:
                        currentdatesplit = currentdate.split('"Version":')
                        currentdate = currentdatesplit[1].strip()
                if '>' in diffline:
                    benchmark = benchmark + diffline
                    benchmarksplits = benchmark.split('>')
                    benchmarkdate = benchmarksplits[1].strip()
                    if len(benchmarkdate) > 0 and '"Value":' in benchmarkdate:
                        benchmarkdatesplit = benchmarkdate.split('"Value":')
                        benchmarkdate = benchmarkdatesplit[1].strip()
                    if len(benchmarkdate) > 0 and '"Version":' in benchmarkdate:
                        benchmarkdatesplit = benchmarkdate.split('"Version":')
                        benchmarkdate = benchmarkdatesplit[1].strip()

            curdate, curflag = DateParser().parse(currentdate)
            benchdate, benflag = DateParser().parse(benchmarkdate)

            if curflag == [True, True, True] and benflag == [True, True, True]:
                continue
            else:
                finaldiff.append(diff)

        return dataobject(finaldiff, nofile)

    def readfiles(self):
        filepath=self.autocaseresultpath
        autocasebenchmarkdiff=search(filepath,'diff_file_benchmark')
        autocasediffbenchmarkresult=self.parsefiles(autocasebenchmarkdiff[0])
        autocaseselfdiff=filepath + os.path.sep + 'diff_file_self'
        autocasediffselfresult= self.parsefiles(autocaseselfdiff)
        result = []
        result.append(autocasediffbenchmarkresult)
        result.append(autocasediffselfresult)
        return result



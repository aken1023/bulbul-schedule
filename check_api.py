import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import json, urllib.request

r = urllib.request.urlopen('http://localhost:5000/api/salary/calculate?year=2026&month=2')
data = json.loads(r.read())['data']

excel = {
    '施秉綬': {'D':16,'Dr':3068,'N':5,'Nr':3168,'extra':2000,'deduct':3346,'pay':63582},
    '陳惠美': {'D':4,'Dr':3058,'N':3,'Nr':3168,'extra':0,'deduct':0,'pay':21736},
    '王明珠': {'D':13,'Dr':3068,'N':0,'Nr':2900,'extra':0,'deduct':0,'pay':39884},
    '何金艷': {'D':17,'Dr':3068,'N':0,'Nr':2900,'extra':0,'deduct':30,'pay':52126},
    '全蕙萍': {'D':16,'Dr':3068,'N':0,'Nr':3168,'extra':8904,'deduct':1196,'pay':56796},
    '衡怡紅': {'D':7,'Dr':2968,'N':3,'Nr':3168,'extra':0,'deduct':0,'pay':30280},
    '藍茂禎': {'D':3,'Dr':3068,'N':20,'Nr':3168,'extra':0,'deduct':30,'pay':72534},
    '林玉燕': {'D':0,'Dr':3068,'N':10,'Nr':3168,'extra':0,'deduct':30,'pay':31650},
    '邱素蘭': {'D':0,'Dr':3068,'N':15,'Nr':3168,'extra':0,'deduct':0,'pay':47520},
    '陳濬宏': {'D':1,'Dr':3068,'N':0,'Nr':3168,'extra':0,'deduct':0,'pay':3068},
    '林淑真': {'D':15,'Dr':2900,'N':0,'Nr':0,'extra':0,'deduct':30,'pay':43470},
    '歐南蘭': {'D':15,'Dr':2700,'N':0,'Nr':0,'extra':1350,'deduct':0,'pay':41850},
    '沈佳榕': {'D':13,'Dr':2700,'N':0,'Nr':0,'extra':600,'deduct':30,'pay':35670},
    '鍾宏枝': {'D':4,'Dr':2600,'N':0,'Nr':0,'extra':0,'deduct':0,'pay':10400},
    '張瀞文': {'D':5,'Dr':2700,'N':0,'Nr':0,'extra':6000,'deduct':30,'pay':19470},
    '蔡詠安': {'D':6,'Dr':2700,'N':18,'Nr':3000,'extra':0,'deduct':0,'pay':70200},
    '王柏欣': {'D':14,'Dr':2700,'N':4,'Nr':3000,'extra':3668,'deduct':3772,'pay':49696},
    '楊張彗': {'D':0,'Dr':2700,'N':23,'Nr':3000,'extra':3000,'deduct':9581,'pay':62419},
    '陳倩儀': {'D':4,'Dr':2600,'N':2,'Nr':2900,'extra':600,'deduct':0,'pay':16800},
}

print(f"{'員工':6s} | {'D天':>3s} {'D價':>5s} | {'N天':>3s} {'N價':>5s} | {'加項':>6s} | {'扣項':>6s} | {'實領':>7s} | 狀態")
print('-' * 80)
for s in data:
    name = s['employee']['name']
    ex = excel.get(name)
    if not ex:
        continue
    issues = []
    if s['day_shifts'] != ex['D']:
        issues.append(f"D天:{s['day_shifts']}vs{ex['D']}")
    if s['day_rate'] != ex['Dr']:
        issues.append(f"D價:{s['day_rate']}vs{ex['Dr']}")
    if s['night_shifts'] != ex['N']:
        issues.append(f"N天:{s['night_shifts']}vs{ex['N']}")
    if s['night_rate'] != ex['Nr']:
        issues.append(f"N價:{s['night_rate']}vs{ex['Nr']}")
    if s['extra_earn_total'] != ex['extra']:
        issues.append(f"加:{s['extra_earn_total']}vs{ex['extra']}")
    if s['deduct_total'] != ex['deduct']:
        issues.append(f"扣:{s['deduct_total']}vs{ex['deduct']}")
    if s['actual_pay'] != ex['pay']:
        issues.append(f"實領:{s['actual_pay']}vs{ex['pay']}")
    status = 'OK' if not issues else ' '.join(issues)
    print(f"{name:6s} | {s['day_shifts']:>3d} {s['day_rate']:>5,} | {s['night_shifts']:>3d} {s['night_rate']:>5,} | {s['extra_earn_total']:>6,} | {s['deduct_total']:>6,} | {s['actual_pay']:>7,} | {status}")

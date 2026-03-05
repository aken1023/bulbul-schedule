import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import json
from datetime import date
from app import app, db
from models import Employee, SalaryRecord

with app.app_context():
    emps = {e.name: e for e in Employee.query.filter_by(is_active=True).all()}

    SalaryRecord.query.filter_by(year=2026, month=2).delete()
    db.session.commit()

    records = [
        # === 雙和 ===
        # 施秉綬: D=16x3068=49088, N=5x3168=15840, L=2000, 總計=66928, 應總扣=3346, 實領=63582
        dict(name='施秉綬', day_shifts=16, night_shifts=5, day_rate=3068, night_rate=3168,
            extra_earnings=[{'name': '加班(L)', 'amount': 2000}],
            advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[{'name': '勞健保等', 'amount': 3346}], note=''),

        # 陳惠美: D=4x3058=12232, N=3x3168=9504, 總計=21736
        dict(name='陳惠美', day_shifts=4, night_shifts=3, day_rate=3058, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),

        # 王明珠: D=13x3068=39884
        dict(name='王明珠', day_shifts=13, night_shifts=0, day_rate=3068, night_rate=2900,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),

        # 何金艷: D=17x3068=52156, 扣轉帳費30
        dict(name='何金艷', day_shifts=17, night_shifts=0, day_rate=3068, night_rate=2900,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note=''),

        # 全蕙萍: 新人3天x2968=8904 + 正常16天x3068=49088 = 57992, 扣勞健1196
        dict(name='全蕙萍', day_shifts=16, night_shifts=0, day_rate=3068, night_rate=3168,
            extra_earnings=[{'name': '新人期D班3天x2968', 'amount': 8904}],
            advance_pay=0, health_insurance=1196, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note='新人3天2968+正常16天3068'),

        # 衡怡紅: 新人D=7天x2968=20776, N=3天x3168=9504, 總計=30280
        dict(name='衡怡紅', day_shifts=7, night_shifts=3, day_rate=2968, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note='環介/D30起為3068'),

        # 藍茂禎: D=3x3068=9204, N=20x3168=63360, 扣轉帳費30
        dict(name='藍茂禎', day_shifts=3, night_shifts=20, day_rate=3068, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note=''),

        # 林玉燕: N=10天x3168=31680, 扣轉帳費30
        dict(name='林玉燕', day_shifts=0, night_shifts=10, day_rate=3068, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note=''),

        # 邱素蘭: N=15x3168=47520
        dict(name='邱素蘭', day_shifts=0, night_shifts=15, day_rate=3068, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),

        # 陳濬宏: 新人月底, 實領=0
        dict(name='陳濬宏', day_shifts=1, night_shifts=0, day_rate=3068, night_rate=3168,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note='115.03起'),

        # === 北醫 ===
        # 林淑真: 15天x2900=43500, 扣轉帳30
        dict(name='林淑真', day_shifts=15, night_shifts=0, day_rate=2900, night_rate=0,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note=''),

        # 歐南蘭: 15天x2700=40500, 6hr加班=1350
        dict(name='歐南蘭', day_shifts=15, night_shifts=0, day_rate=2700, night_rate=0,
            extra_earnings=[{'name': '6hr加班', 'amount': 1350}],
            advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),

        # 沈佳榕: 13天x2700=35100, 高配比=600, 扣轉帳30
        dict(name='沈佳榕', day_shifts=13, night_shifts=0, day_rate=2700, night_rate=0,
            extra_earnings=[{'name': '高配比', 'amount': 600}],
            advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note=''),

        # 鍾宏枝: 4天x2600=10400
        dict(name='鍾宏枝', day_shifts=4, night_shifts=0, day_rate=2600, night_rate=0,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),

        # 張瀞文: 5天x2700=13500 + 何信2天x3000=6000, 扣轉帳30
        dict(name='張瀞文', day_shifts=5, night_shifts=0, day_rate=2700, night_rate=0,
            extra_earnings=[{'name': '何信2天x3000', 'amount': 6000}],
            advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[], note='北醫5天+何信2天'),

        # 蔡詠安: 北醫+和信 (no salary data in Excel)
        dict(name='蔡詠安', day_shifts=6, night_shifts=18, day_rate=2700, night_rate=3000,
            extra_earnings=[], advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note='北醫+和信'),

        # 王柏欣: D=14x2700=37800, N=4x3000=12000, 雙和1天=3068, 高配比+2hr加班=600
        #          扣勞健1915+其他1827+轉帳30=3772
        dict(name='王柏欣', day_shifts=14, night_shifts=4, day_rate=2700, night_rate=3000,
            extra_earnings=[{'name': '雙和1天x3068', 'amount': 3068}, {'name': '高配比+2hr加班', 'amount': 600}],
            advance_pay=0, health_insurance=1915, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[{'name': '其他扣', 'amount': 1827}], note='北醫+雙和'),

        # 楊張彗: N=23x3000=69000, 高配比3天=3000, 扣勞健7201+其他2350+轉帳30
        dict(name='楊張彗', day_shifts=0, night_shifts=23, day_rate=2700, night_rate=3000,
            extra_earnings=[{'name': '高配比3天', 'amount': 3000}],
            advance_pay=0, health_insurance=7201, welfare_fund=0, leave_deduct=0, transfer_fee=30,
            extra_deductions=[{'name': '其他扣(減5%/超25000)', 'amount': 2350}], note=''),

        # 陳倩儀: D=4x2600=10400, N=2x2900=5800, 高配比=600
        dict(name='陳倩儀', day_shifts=4, night_shifts=2, day_rate=2600, night_rate=2900,
            extra_earnings=[{'name': '高配比', 'amount': 600}],
            advance_pay=0, health_insurance=0, welfare_fund=0, leave_deduct=0, transfer_fee=0,
            extra_deductions=[], note=''),
    ]

    for r in records:
        emp = emps.get(r['name'])
        if not emp:
            print(f'WARNING: {r["name"]} not found')
            continue

        rec = SalaryRecord(
            employee_id=emp.id, year=2026, month=2,
            day_shifts=r['day_shifts'], night_shifts=r['night_shifts'],
            day_rate=r['day_rate'], night_rate=r['night_rate'],
            extra_earnings=json.dumps(r['extra_earnings'], ensure_ascii=False),
            advance_pay=r['advance_pay'], health_insurance=r['health_insurance'],
            welfare_fund=r['welfare_fund'], leave_deduct=r['leave_deduct'],
            transfer_fee=r['transfer_fee'],
            extra_deductions=json.dumps(r['extra_deductions'], ensure_ascii=False),
            note=r['note'],
        )
        db.session.add(rec)

        gross = r['day_shifts'] * r['day_rate'] + r['night_shifts'] * r['night_rate']
        extra_e = sum(e['amount'] for e in r['extra_earnings'])
        total_earn = gross + extra_e
        deducts = r['advance_pay'] + r['health_insurance'] + r['welfare_fund'] + r['leave_deduct'] + r['transfer_fee']
        extra_d = sum(d['amount'] for d in r['extra_deductions'])
        total_deduct = deducts + extra_d
        actual = total_earn - total_deduct
        print(f'{r["name"]:6s} | earn={total_earn:>6,} | deduct={total_deduct:>5,} | actual={actual:>6,}')

    db.session.commit()
    print(f'\nSaved {len(records)} salary records')

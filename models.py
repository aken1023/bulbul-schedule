from flask_sqlalchemy import SQLAlchemy
from datetime import date

db = SQLAlchemy()


class Location(db.Model):
    __tablename__ = 'locations'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    max_consecutive_days = db.Column(db.Integer, default=6)  # 最大連續上班天數

    def to_dict(self):
        return {
            'id': self.id, 'name': self.name, 'is_active': self.is_active,
            'max_consecutive_days': self.max_consecutive_days or 6,
        }


class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    emp_code = db.Column(db.String(20))
    start_date = db.Column(db.Date)
    location = db.Column(db.String(20))  # 雙和 / 北醫
    job_title = db.Column(db.String(30), default='照服員')
    bank_account = db.Column(db.String(60))
    day_rate = db.Column(db.Integer, default=3068)
    night_rate = db.Column(db.Integer, default=3168)
    is_active = db.Column(db.Boolean, default=True)

    schedules = db.relationship('Schedule', backref='employee', lazy='dynamic')

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'emp_code': self.emp_code or '',
            'start_date': self.start_date.isoformat() if self.start_date else '',
            'location': self.location or '',
            'job_title': self.job_title or '照服員',
            'bank_account': self.bank_account or '',
            'day_rate': self.day_rate,
            'night_rate': self.night_rate,
            'is_active': self.is_active,
        }


class Schedule(db.Model):
    __tablename__ = 'schedules'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    shift = db.Column(db.String(10), default='OFF')  # D, N, OFF

    __table_args__ = (db.UniqueConstraint('employee_id', 'date'),)


class SalaryRecord(db.Model):
    __tablename__ = 'salary_records'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    month = db.Column(db.Integer, nullable=False)

    day_shifts = db.Column(db.Integer, default=0)
    night_shifts = db.Column(db.Integer, default=0)
    day_rate = db.Column(db.Integer, default=0)
    night_rate = db.Column(db.Integer, default=0)

    # 發金額 (earnings)
    extra_earnings = db.Column(db.Text, default='[]')  # JSON: [{"name":"體檢費","amount":1510}]

    # 扣金額 (deductions)
    advance_pay = db.Column(db.Integer, default=0)      # 提早撥款本月薪資
    health_insurance = db.Column(db.Integer, default=0)  # 健保費眷屬
    welfare_fund = db.Column(db.Integer, default=0)      # 福利金
    leave_deduct = db.Column(db.Integer, default=0)      # 請假
    transfer_fee = db.Column(db.Integer, default=0)      # 匯費
    extra_deductions = db.Column(db.Text, default='[]')  # JSON: [{"name":"其他","amount":100}]

    note = db.Column(db.Text, default='')
    payment_date = db.Column(db.Date)

    employee = db.relationship('Employee', backref='salary_records')

    __table_args__ = (db.UniqueConstraint('employee_id', 'year', 'month'),)


class EmployeePreference(db.Model):
    __tablename__ = 'employee_preferences'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False, unique=True)
    allowed_locations = db.Column(db.Text, default='[]')   # JSON list: ["雙和","北醫"]
    allowed_shifts = db.Column(db.String(10), default='both')  # "D", "N", "both"
    min_days = db.Column(db.Integer, default=0)   # 月最低上班天數
    max_days = db.Column(db.Integer, default=31)  # 月最高上班天數
    # 工作模式: "none"(不指定), "weekdays"(指定星期幾), "pattern"(做X休Y)
    schedule_mode = db.Column(db.String(20), default='none')
    # 指定上班星期幾 JSON list: [0,1,2,3,4] (0=一,6=日)
    work_weekdays = db.Column(db.Text, default='[]')
    # 做X休Y模式
    pattern_work = db.Column(db.Integer, default=5)  # 連續工作天數
    pattern_off = db.Column(db.Integer, default=2)    # 連續休息天數

    employee = db.relationship('Employee', backref=db.backref('preference', uselist=False))

    def to_dict(self):
        import json
        return {
            'id': self.id,
            'employee_id': self.employee_id,
            'allowed_locations': json.loads(self.allowed_locations) if self.allowed_locations else [],
            'allowed_shifts': self.allowed_shifts or 'both',
            'min_days': self.min_days,
            'max_days': self.max_days,
            'schedule_mode': self.schedule_mode or 'none',
            'work_weekdays': json.loads(self.work_weekdays) if self.work_weekdays else [],
            'pattern_work': self.pattern_work or 5,
            'pattern_off': self.pattern_off or 2,
        }


class EmployeeTimeOff(db.Model):
    """員工預定排休（指定不想上班的日期）"""
    __tablename__ = 'employee_time_off'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    reason = db.Column(db.String(100), default='')

    __table_args__ = (db.UniqueConstraint('employee_id', 'date'),)

    employee = db.relationship('Employee', backref='time_offs')

    def to_dict(self):
        return {
            'id': self.id,
            'employee_id': self.employee_id,
            'date': self.date.isoformat(),
            'reason': self.reason or '',
        }


class EmployeeRateHistory(db.Model):
    """員工薪資費率歷史：自某日起適用的日班/夜班費率"""
    __tablename__ = 'employee_rate_history'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    effective_date = db.Column(db.Date, nullable=False)  # 生效日期
    day_rate = db.Column(db.Integer, nullable=False)
    night_rate = db.Column(db.Integer, nullable=False)

    employee = db.relationship('Employee', backref=db.backref('rate_history', lazy='dynamic'))

    __table_args__ = (db.UniqueConstraint('employee_id', 'effective_date'),)

    def to_dict(self):
        return {
            'id': self.id,
            'employee_id': self.employee_id,
            'effective_date': self.effective_date.isoformat(),
            'day_rate': self.day_rate,
            'night_rate': self.night_rate,
        }


class StaffingRequirement(db.Model):
    __tablename__ = 'staffing_requirements'
    id = db.Column(db.Integer, primary_key=True)
    location = db.Column(db.String(50), nullable=False)
    shift = db.Column(db.String(10), nullable=False)  # D or N
    count = db.Column(db.Integer, default=0)

    __table_args__ = (db.UniqueConstraint('location', 'shift'),)

    def to_dict(self):
        return {
            'id': self.id,
            'location': self.location,
            'shift': self.shift,
            'count': self.count,
        }

"""
自動排班演算法 - 貪婪約束滿足

硬約束：
1. N→D 禁止（夜班隔天不能接日班，需至少1天OFF）
2. 醫院最大連續上班天數限制
3. 員工班別限制（僅日班/僅夜班/皆可）
4. 員工院區限制（含跨院設定）
5. 月最高/最低上班天數
6. 排班模式：指定星期 / 做X休Y / 不指定
7. 預定排休日
8. 每週至少休息1天（每連續7天內至少1天OFF）

軟約束：均衡班次分配（工作天數少者優先）
"""
import calendar
from datetime import date


def auto_generate(year, month, employees, preferences, staffing_reqs, locations, loc_max_consecutive, time_off_dates=None):
    """
    Args:
        year, month: int
        employees: list of Employee objects
        preferences: dict {emp_id: pref_dict}
        staffing_reqs: dict {(location, shift): count}
        locations: list of location name strings
        loc_max_consecutive: dict {location_name: max_days}
        time_off_dates: dict {emp_id: set of day ints} - 員工預定排休日
    Returns:
        dict {emp_id: {day(int): shift_str}}
    """
    if time_off_dates is None:
        time_off_dates = {}
    _, days_in_month = calendar.monthrange(year, month)

    # Global max consecutive (fallback if no location-specific limit)
    global_max_consec = min(loc_max_consecutive.values()) if loc_max_consecutive else 6

    result = {emp.id: {d: 'OFF' for d in range(1, days_in_month + 1)} for emp in employees}
    work_counts = {emp.id: 0 for emp in employees}
    shortages = []  # 人力不足記錄: [{day, location, shift, required, assigned}]

    # Pre-compute which days each employee is AVAILABLE based on schedule_mode
    available = {}
    for emp in employees:
        pref = preferences.get(emp.id, {})
        mode = pref.get('schedule_mode', 'none')

        if mode == 'weekdays':
            work_wd = set(pref.get('work_weekdays', []))
            avail_days = set()
            for d in range(1, days_in_month + 1):
                dt = date(year, month, d)
                # Python weekday: 0=Mon. Our UI: 0=一(Mon), so they match
                if dt.weekday() in work_wd:
                    avail_days.add(d)
        elif mode == 'pattern':
            pw = pref.get('pattern_work', 5)
            po = pref.get('pattern_off', 2)
            cycle = pw + po
            avail_days = set()
            for d in range(1, days_in_month + 1):
                if ((d - 1) % cycle) < pw:
                    avail_days.add(d)
        else:
            avail_days = set(range(1, days_in_month + 1))

        available[emp.id] = avail_days

    def get_consecutive_before(emp_id, day):
        """Count consecutive work days ending at day-1 (before current day)."""
        count = 0
        for d in range(day - 1, 0, -1):
            if result[emp_id][d] != 'OFF':
                count += 1
            else:
                break
        return count

    def would_violate_weekly_rest(emp_id, day):
        """Check if assigning work on this day would cause 7+ consecutive work days."""
        # Count consecutive work days before this day
        before = get_consecutive_before(emp_id, day)
        # Count consecutive work days after this day (already assigned)
        after = 0
        for d in range(day + 1, min(day + 7, days_in_month + 1)):
            if result[emp_id][d] != 'OFF':
                after += 1
            else:
                break
        # Total consecutive if we assign this day
        total = before + 1 + after
        return total >= 7

    def can_assign(emp_id, day, shift, loc, pref, max_consec):
        """Check ALL constraints for assigning a shift."""
        # 1. Available day (schedule mode)
        if day not in available[emp_id]:
            return False

        # 2. Time off (預定排休)
        if day in time_off_dates.get(emp_id, set()):
            return False

        # 3. Location constraint
        allowed_locs = pref.get('allowed_locations', [])
        if allowed_locs and loc not in allowed_locs:
            return False
        if not allowed_locs:
            emp = next(e for e in employees if e.id == emp_id)
            if emp.location != loc:
                return False

        # 4. Shift constraint
        allowed_shifts = pref.get('allowed_shifts', 'both')
        if allowed_shifts == 'D' and shift != 'D':
            return False
        if allowed_shifts == 'N' and shift != 'N':
            return False

        # 5. Max monthly days
        max_days = pref.get('max_days', 31)
        if work_counts[emp_id] >= max_days:
            return False

        # 6. N→D constraint (夜班隔天不能接日班)
        if shift == 'D' and day > 1 and result[emp_id][day - 1] == 'N':
            return False

        # 7. Max consecutive days (醫院規定)
        consec = get_consecutive_before(emp_id, day)
        if consec >= max_consec:
            return False

        # 8. Weekly rest (每7天至少1天OFF)
        if would_violate_weekly_rest(emp_id, day):
            return False

        return True

    # === Main scheduling loop: for each day, fill demands ===
    for day in range(1, days_in_month + 1):
        assigned_today = set()

        demands = []
        for loc in locations:
            for shift in ['D', 'N']:
                req = staffing_reqs.get((loc, shift), 0)
                if req > 0:
                    demands.append((loc, shift, req))

        for loc, shift, req_count in demands:
            max_consec = loc_max_consecutive.get(loc, 6)
            candidates = []

            for emp in employees:
                if emp.id in assigned_today:
                    continue
                pref = preferences.get(emp.id, {})
                if can_assign(emp.id, day, shift, loc, pref, max_consec):
                    candidates.append(emp)

            # Sort by work count ascending for fairness
            candidates.sort(key=lambda e: work_counts[e.id])

            filled = 0
            for emp in candidates:
                if filled >= req_count:
                    break
                result[emp.id][day] = shift
                assigned_today.add(emp.id)
                work_counts[emp.id] += 1
                filled += 1

            if filled < req_count:
                shortages.append({
                    'day': day,
                    'location': loc,
                    'shift': shift,
                    'required': req_count,
                    'assigned': filled,
                    'short': req_count - filled,
                })

    # === Post-processing: enforce min_days ===
    for emp in employees:
        pref = preferences.get(emp.id, {})
        min_days = pref.get('min_days', 0)
        if work_counts[emp.id] >= min_days:
            continue

        allowed_shifts = pref.get('allowed_shifts', 'both')
        needed = min_days - work_counts[emp.id]

        for day in range(1, days_in_month + 1):
            if needed <= 0:
                break
            if result[emp.id][day] != 'OFF':
                continue

            # Determine shift type
            shift = 'D' if allowed_shifts != 'N' else 'N'

            # Use employee's own location or first allowed location
            allowed_locs = pref.get('allowed_locations', [])
            loc = allowed_locs[0] if allowed_locs else emp.location
            max_consec = loc_max_consecutive.get(loc, global_max_consec)

            # Run full constraint check (reuse can_assign with a dummy loc match)
            # Check available day
            if day not in available[emp.id]:
                continue
            # Check time off
            if day in time_off_dates.get(emp.id, set()):
                continue
            # Check N→D
            if shift == 'D' and day > 1 and result[emp.id][day - 1] == 'N':
                continue
            # Check D→N for next day (if we assign N, next day can't be D already)
            if shift == 'N' and day < days_in_month and result[emp.id][day + 1] == 'D':
                continue
            # Check max consecutive
            consec = get_consecutive_before(emp.id, day)
            if consec >= max_consec:
                continue
            # Check weekly rest
            if would_violate_weekly_rest(emp.id, day):
                continue

            result[emp.id][day] = shift
            work_counts[emp.id] += 1
            needed -= 1

    return result, shortages

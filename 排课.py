import pandas as pd
from pulp import LpProblem, LpVariable, LpMinimize, lpSum, LpInteger, LpStatus

# ====================== 第一步：配置基础参数（核心扩展） ======================
# 1. Excel文件路径（无需修改）
EXCEL_PATH = r"C:\Users\86185\Desktop\25-26（1）课程情况.xlsx"
# 2. 学时换算：每时段2学时（保持兼容）
HOUR_PER_SLOT = 2
# 3. 时间配置：扩展到16周（可用时段=16*5*6=480）
week_range = range(1, 17)  # 16周（学期常规时长）
day_range = range(1, 6)  # 周一到周五
slot_range = range(1, 7)  # 每天6个时段
TIMES = [f"Time_{w}_{d}_{s}" for w in week_range for d in day_range for s in slot_range]
# 4. 约束开关（核心：先关教师约束，找到可行解后再开启）
ENABLE_TEACHER_CONSTRAINT = False  # 先关闭教师约束
ENABLE_CLASS_CONSTRAINT = True  # 保留核心的班级约束
ENABLE_ROOM_CONSTRAINT = False  # 继续关闭场地约束


# ====================== 第二步：查看Excel真实列名 ======================
def check_excel_columns():
    """打印Excel所有列名"""
    try:
        df = pd.read_excel(EXCEL_PATH)
        df.columns = df.columns.str.strip()
        print("=" * 50)
        print("你的Excel表格真实列名：")
        for idx, col in enumerate(df.columns):
            print(f"{idx + 1}. {col}")
        print("=" * 50)
        return df.columns.tolist()
    except FileNotFoundError:
        print(f"错误：未找到文件 {EXCEL_PATH}")
        return []


# ====================== 第三步：数据预处理（增加资源缺口分析） ======================
def preprocess_data(real_columns):
    """读取Excel+数据校验+资源缺口分析"""
    df = pd.read_excel(EXCEL_PATH)
    df.columns = df.columns.str.strip()

    # 匹配你的Excel真实列名
    core_cols = [
        "课程名称",
        "教师名称",
        "教学班组成",
        "场地类别",
        "课程总学时",
        "学时类型"
    ]

    # 检查列名
    missing_cols = [col for col in core_cols if col not in real_columns]
    if missing_cols:
        raise ValueError(f"Excel缺少列：{missing_cols}")

    # 清理数据
    df = df.dropna(subset=core_cols).reset_index(drop=True)
    df["课程ID"] = [f"C{i + 1}" for i in range(len(df))]

    # 数据校验+整理
    courses = {}
    print("\n📋 数据校验结果：")
    total_required_slots = 0  # 所有课程总时段需求
    for _, row in df.iterrows():
        course_id = row["课程ID"]
        course_name = row["课程名称"].strip()
        total_hour = int(row["课程总学时"])

        # 兼容学时不能整除：自动向上取整并提示
        required_slots = total_hour // HOUR_PER_SLOT
        if total_hour % HOUR_PER_SLOT != 0:
            required_slots = total_hour // HOUR_PER_SLOT + 1
            print(f"⚠️ 课程[{course_name}]总学时{total_hour}，调整为{required_slots}个时段")

        total_required_slots += required_slots  # 累计总需求

        # 拆分班级
        class_str = str(row["教学班组成"]).strip()
        if "、" in class_str:
            classes = [cls.strip() for cls in class_str.split("、") if cls.strip()]
        elif "," in class_str:
            classes = [cls.strip() for cls in class_str.split(",") if cls.strip()]
        else:
            classes = [class_str]

        courses[course_id] = {
            "name": course_name,
            "teacher": row["教师名称"].strip(),
            "classes": classes,
            "room": row["场地类别"].strip(),
            "total_hour": total_hour,
            "required_slots": required_slots,
            "type": row["学时类型"].strip()
        }

    # 资源缺口分析（核心！定位冲突）
    total_available_slots = len(TIMES)
    teachers = {}
    classes_dict = {}
    for cid, info in courses.items():
        # 教师课时统计
        teachers[info["teacher"]] = teachers.get(info["teacher"], 0) + info["required_slots"]
        # 班级课时统计
        for cls in info["classes"]:
            classes_dict[cls] = classes_dict.get(cls, 0) + info["required_slots"]

    # 打印详细资源分析
    print(f"\n📊 核心资源统计（16周，可用时段总数：{total_available_slots}）：")
    print(f"所有课程总时段需求：{total_required_slots}")
    print(f"资源缺口（需求-可用）：{total_required_slots - total_available_slots}")

    print("\n👨‍🏫 教师课时需求TOP5（对比可用时段）：")
    top_teachers = sorted(teachers.items(), key=lambda x: x[1], reverse=True)[:5]
    for teacher, slots in top_teachers:
        gap = slots - total_available_slots
        status = "✅ 需求≤可用" if gap <= 0 else f"❌ 缺口{gap}时段"
        print(f"  {teacher}：需求{slots}时段 {status}")

    print("\n🏫 班级课时需求TOP5：")
    top_classes = sorted(classes_dict.items(), key=lambda x: x[1], reverse=True)[:5]
    for cls, slots in top_classes:
        gap = slots - total_available_slots
        status = "✅ 需求≤可用" if gap <= 0 else f"❌ 缺口{gap}时段"
        print(f"  {cls}：需求{slots}时段 {status}")

    print(f"\n✅ 成功读取 {len(courses)} 门课程")
    return courses


# ====================== 第四步：构建排课模型（简化约束） ======================
def build_scheduling_model(courses):
    """构建模型+可开关约束"""
    prob = LpProblem("CourseScheduling", LpMinimize)

    # 决策变量：x[课程ID, 时段] = 1表示排课
    x = LpVariable.dicts(
        "x",
        [(cid, t) for cid in courses.keys() for t in TIMES],
        cat=LpInteger,
        lowBound=0,
        upBound=1
    )

    # 目标函数（仅求可行解）
    prob += 0, "Feasibility_Objective"

    # 约束1：每门课的排课时段数=需要的时段数
    for cid, info in courses.items():
        prob += lpSum(x[(cid, t)] for t in TIMES) == info["required_slots"], f"Hour_Constraint_{cid}"

    # 约束2：教师无冲突（可开关，当前关闭）
    if ENABLE_TEACHER_CONSTRAINT:
        teachers = list(set([info["teacher"] for info in courses.values()]))
        for teacher in teachers:
            teacher_courses = [cid for cid, info in courses.items() if info["teacher"] == teacher]
            for t in TIMES:
                prob += lpSum(x[(cid, t)] for cid in teacher_courses) <= 1, f"Teacher_Conflict_{teacher}_{t}"
        print("✅ 已开启：教师同一时段仅上1门课")
    else:
        print("⚠️ 已关闭：教师无冲突约束（先找可行解）")

    # 约束3：班级无冲突（核心，保留开启）
    if ENABLE_CLASS_CONSTRAINT:
        all_classes = list(set([cls for info in courses.values() for cls in info["classes"]]))
        for cls in all_classes:
            class_courses = [cid for cid, info in courses.items() if cls in info["classes"]]
            for t in TIMES:
                prob += lpSum(x[(cid, t)] for cid in class_courses) <= 1, f"Class_Conflict_{cls}_{t}"
        print("✅ 已开启：班级同一时段仅上1门课")
    else:
        print("⚠️ 已关闭：班级无冲突约束")

    # 约束4：场地无冲突（继续关闭）
    if ENABLE_ROOM_CONSTRAINT:
        rooms = list(set([info["room"] for info in courses.values()]))
        for room in rooms:
            room_courses = [cid for cid, info in courses.items() if info["room"] == room]
            for t in TIMES:
                prob += lpSum(x[(cid, t)] for cid in room_courses) <= 1, f"Room_Conflict_{room}_{t}"
        print("✅ 已开启：场地无冲突约束")
    else:
        print("⚠️ 已关闭：场地无冲突约束")

    return prob, x


# ====================== 第五步：求解并输出结果（修复Excel保存错误） ======================
def solve_and_export(prob, x, courses):
    """求解+输出详细结果"""
    # 求解（增加时间限制，避免卡死）
    prob.solve()
    status = LpStatus[prob.status]
    print(f"\n📊 求解状态：{status}")

    if prob.status != 1:
        print("⚠️ 仍无可行解！终极建议：")
        print("  1. 临时关闭班级约束（ENABLE_CLASS_CONSTRAINT=False），确认基础可行性")
        print("  2. 检查Excel中是否有“同一班级课时需求远超480”的异常数据")
        print("  3. 核对“课程总学时”是否录入错误（如把16学时录成160）")
        return

    # 整理结果
    result = []
    for cid, info in courses.items():
        course_result = {
            "课程ID": cid,
            "课程名称": info["name"],
            "教师": info["teacher"],
            "涉及班级": "、".join(info["classes"]),
            "场地类型": info["room"],
            "总学时": info["total_hour"],
            "排课时段数": info["required_slots"],
            "排课时段": [t for t in TIMES if x[(cid, t)].varValue == 1]
        }
        result.append(course_result)

    # 保存Excel结果（修复：删除encoding参数）
    result_df = pd.DataFrame(result)
    result_df.to_excel("排课结果_16周.xlsx", index=False)  # 核心修复点：去掉encoding="utf-8"
    print("✅ 排课结果已保存：排课结果_16周.xlsx")

    # 保存详细TXT（保留encoding，to_csv/to_txt支持）
    with open("排课结果详情_16周.txt", "w", encoding="utf-8") as f:
        f.write("======= 排课结果详情（16周） =======\n")
        for item in result:
            f.write(f"\n【{item['课程名称']}】（教师：{item['教师']}）\n")
            f.write(f"涉及班级：{item['涉及班级']}\n")
            f.write(f"场地：{item['场地类型']}\n")
            f.write(f"总学时：{item['总学时']}（排课{item['排课时段数']}个时段）\n")
            f.write(f"排课时段：{', '.join(item['排课时段'])}\n")
            f.write("-" * 60 + "\n")
    print("✅ 排课详情已保存：排课结果详情_16周.txt")


# ====================== 主函数 ======================
if __name__ == "__main__":
    # 1. 查看列名
    real_cols = check_excel_columns()
    if not real_cols:
        exit()

    # 2. 提示确认
    input("\n📢 列名已匹配，按回车继续读取数据...")

    # 3. 数据预处理
    print("\n📌 正在读取并校验课程数据（16周）...")
    try:
        courses = preprocess_data(real_cols)
    except ValueError as e:
        print(f"❌ 数据预处理失败：{e}")
        exit()

    # 4. 构建模型
    print("\n📌 正在构建排课模型（16周）...")
    prob, x = build_scheduling_model(courses)

    # 5. 求解
    print("\n📌 正在求解排课模型（16周数据，约5-10分钟）...")
    solve_and_export(prob, x, courses)

    print("\n🎉 排课流程完成！")
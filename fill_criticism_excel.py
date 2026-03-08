import argparse
import datetime as dt
import random
from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class Pair:
    problem: str
    suggest: str


PAIRS: list[Pair] = [
    Pair(
        problem="政治理论学习有时偏碎片化，学用结合不够紧密，理论对科研攻关的指导性还需增强",
        suggest="建议制定年度/季度理论学习计划，围绕智慧农业与智能农机重点任务开展专题研学，并在项目立项、阶段评审中主动用理论校准方向",
    ),
    Pair(
        problem="在科研项目推进中对关键节点的统筹和风险预判不够，前瞻性谋划与资源协调还有提升空间",
        suggest="建议强化里程碑管理和风险清单机制，重要节点提前评估人力/物料/测试资源，遇到跨部门依赖及时沟通协调，提升项目交付确定性",
    ),
    Pair(
        problem="与团队成员的沟通协作有时不够主动，技术方案共识与信息同步存在滞后",
        suggest="建议定期组织技术评审与经验复盘，关键方案形成可共享的文档与结论沉淀，主动与相关同事对齐接口与边界，提升协作效率",
    ),
    Pair(
        problem="对智慧农业前沿技术（感知、定位导航、作业决策、农业大模型等）的系统跟踪还不够，创新点凝练不够突出",
        suggest="建议建立前沿跟踪清单与月度分享机制，聚焦与现有产品/平台结合的可落地点，及时将新方法转化为可验证的原型与指标对比",
    ),
    Pair(
        problem="服务产业需求时对一线应用场景调研不够深入，对痛点与约束条件把握不够细",
        suggest="建议增加到试验场、示范基地和用户现场的调研频次，形成场景画像与需求清单，推动需求闭环到方案、验证与迭代",
    ),
    Pair(
        problem="工作精细化管理还有短板，部分过程资料、实验记录与数据管理不够规范",
        suggest="建议强化过程管理与质量意识，关键实验建立统一模板与版本管理，数据采集与标注形成规范流程，提升复现性与成果可交付性",
    ),
    Pair(
        problem="在成果凝练与落地转化方面投入不够，阶段性成果对外输出与推广力度有待加强",
        suggest="建议把论文/专利/标准/软件著作等成果计划纳入项目里程碑，同时加强与产品、应用团队联动，推动成果形成可用组件与应用示范",
    ),
    Pair(
        problem="面对复杂任务时攻坚韧劲与主动担当还可以更强，遇到难点时有时偏保守",
        suggest="建议增强攻坚意识，针对关键瓶颈设定“问题负责人+周跟踪”机制，敢于尝试多路线并行验证，及时复盘收敛最优路径",
    ),
]


def _normalize_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\r", "").strip()


def _is_placeholder(value: str) -> bool:
    s = value.strip().replace(" ", "")
    return s in {"", "nan", "none", "1.\n2.", "1.\n2.\n"}


def _pick_two_distinct(rng: random.Random) -> tuple[Pair, Pair]:
    first = rng.choice(PAIRS)
    second_pool = [p for p in PAIRS if p != first]
    second = rng.choice(second_pool)
    return first, second


def make_text(name: str, seq: int) -> str:
    # 使用“姓名+序号”做种子，保证每次生成稳定、同时人与人之间差异更大
    seed_src = f"{name}#{seq}"
    seed = int.from_bytes(seed_src.encode("utf-8"), "little", signed=False) % (2**32)
    rng = random.Random(seed)

    p1, p2 = _pick_two_distinct(rng)

    # 输出保持与原表一致的“1./2.”格式，仅保留建议内容
    return (
        f"1.建议：{p1.suggest}。\n"
        f"2.建议：{p2.suggest}。"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input",
        default=r"d:\\桌面\\0305党组织生活会材料\\批评意见表.xlsx",
        help="Input Excel path",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Output Excel path (default: auto-generated in same folder)",
    )
    args = parser.parse_args()

    input_path = args.input
    df = pd.read_excel(input_path)

    target_col = "存在的问题及建议"
    if target_col not in df.columns:
        df[target_col] = ""

    name_col = "提出意见对象" if "提出意见对象" in df.columns else None
    seq_col = "序号" if "序号" in df.columns else None

    normalized = df[target_col].map(_normalize_cell)
    to_fill_mask = normalized.map(_is_placeholder)

    fill_count = int(to_fill_mask.sum())

    def _row_name(i: int) -> str:
        if name_col is None:
            return ""
        v = df.at[i, name_col]
        return "" if pd.isna(v) else str(v).strip()

    def _row_seq(i: int) -> int:
        if seq_col is None:
            return i + 1
        v = df.at[i, seq_col]
        try:
            return int(v)
        except Exception:
            return i + 1

    for i in df.index[to_fill_mask]:
        name = _row_name(i)
        seq = _row_seq(i)
        df.at[i, target_col] = make_text(name=name, seq=seq)

    if args.output:
        out_path = args.output
    else:
        stamp = dt.date.today().strftime("%Y%m%d")
        out_path = input_path.replace(
            ".xlsx", f"_建议版_{stamp}.xlsx"
        )

    df.to_excel(out_path, index=False)
    print(f"filled: {fill_count}")
    print(f"output: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

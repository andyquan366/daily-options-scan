import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

import yfinance as yf
import pandas as pd
import os

# 关闭缓存，强制拉最新数据
os.environ["YFINANCE_NO_CACHE"] = "1"

try:
    from yfinance import Ticker
    Ticker.cache_disable()
except:
    pass


# ============================================================
# ★ 载入数据 ★
# ============================================================
def load_price(ticker):
    df = yf.download(
        ticker,
        start="2020-01-01",
        end=None,
        interval="1d",
        auto_adjust=False,
        progress=False,
        ignore_tz=True
    )

    if df is None or df.empty:
        raise ValueError(f"{ticker} empty")

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.droplevel(list(range(df.columns.nlevels - 1)))

    if df.shape[1] < 3:
        raise ValueError(f"{ticker} 数据列太少: {df.columns}")

    # 强制按列顺序取 High / Low
    df["HIGH"] = df.iloc[:, 1].astype(float)
    df["LOW"]  = df.iloc[:, 2].astype(float)
    
    return df


spy = load_price("SPY")
qqq = load_price("QQQ")

print("SPY 数据日期范围:", spy.index.min(), "~", spy.index.max())
print("QQQ 数据日期范围:", qqq.index.min(), "~", qqq.index.max())
print("="*80)


# ============================================================
# ★★ 事件式回撤（修正天数计算逻辑）★★
# ============================================================
def detect_events(df, threshold, name):

    high = df["HIGH"]
    low  = df["LOW"]
    
    df_index = df.index
    events = []
    
    # 状态变量
    in_dd = False
    event_peak_date = None
    event_peak_price = None
    event_low_date = None
    event_low_price = None
    
    # 【修正 1：引入周期峰值】用于消除重复事件 (493天问题)
    current_cycle_peak_date = df_index[0]

    for i, date in enumerate(df_index):
        cur_low  = float(low.loc[date])
        cur_high = float(high.loc[date])
        
        # 追踪当前周期峰值：只有创出新高，才更新当前周期的峰值日期
        if cur_high >= float(high.loc[current_cycle_peak_date]):
            current_cycle_peak_date = date

        # ------------------- 触发回撤 -------------------
        # 触发条件：不在回撤中 AND (当前最低价 / 当前周期峰值 - 1) <= -threshold
        current_peak_price = float(high.loc[current_cycle_peak_date])
        cur_dd = cur_low / current_peak_price - 1
        
        if not in_dd and cur_dd <= -threshold:
            in_dd = True
            
            # 锁定事件的峰值信息
            event_peak_date = current_cycle_peak_date
            event_peak_price = current_peak_price
            
            # 记录初始低点
            event_low_date = date
            event_low_price = cur_low

        # ------------------- 更新低点 -------------------
        if in_dd and cur_low < event_low_price:
            event_low_price = cur_low
            event_low_date = date
                
        # ------------------- 回撤结束 -------------------
        # 结束条件：正在回撤中 AND (当前最低价 / 事件峰值 > -threshold)
        if in_dd:
            # 使用事件峰值计算 DD
            dd_vs_event_peak = cur_low / event_peak_price - 1
            
            if dd_vs_event_peak > -threshold:
                
                # 【修正 2：使用索引确定前一个交易日】修正 End Date 错误
                if i > 0:
                    end_date = df_index[i - 1]
                else:
                    # 避免在第一天就结束的情况（虽然不可能触发）
                    end_date = event_peak_date 

                max_dd = event_low_price / event_peak_price - 1
                
                # ★★★ 修正天数指标计算逻辑 ★★★
                # 1. 总持续天数 (Peak 到 End)
                days_total = df.loc[event_peak_date:end_date].shape[0]
                # 2. 回撤到谷底天数 (Peak 到 Bottom, 包含两端)
                days_to_bottom = df.loc[event_peak_date:event_low_date].shape[0]
                # 3. 回弹天数 (新逻辑: 总天数 - 回撤天数 = 谷底后到结束的天数)
                #    让 DaysToBottom 包含谷底日，DaysToRecovery 包含谷底后的所有天数。
                days_to_recovery = days_total - days_to_bottom
                # ***********************************

                events.append({
                    "Start": event_peak_date.strftime("%Y-%m-%d"),
                    "End": end_date.strftime("%Y-%m-%d"),
                    "Bottom": event_low_date.strftime("%Y-%m-%d"),
                    "PeakHigh": f"{event_peak_price:.2f}",
                    "BottomLow": f"{event_low_price:.2f}",
                    "MaxDD": f"{max_dd*100:.2f}%",
                    "DaysTotal": days_total,             
                    "DaysToBottom": days_to_bottom,      
                    "DaysToRecovery": days_to_recovery,  
                })

                in_dd = False
                
                # 【修正 1 补充】回撤结束后，将周期峰值标记为当前日期，
                # 阻止在没有新高的情况下再次触发旧峰值的事件。
                current_cycle_peak_date = date

    # ------------------- 尾部事件 -------------------
    # 【修正 3：修复尾部逻辑】只有当 in_dd 为 True (事件在结束时仍在进行中)才记录
    if in_dd: 
        end_date = df_index[-1]
        max_dd = event_low_price / event_peak_price - 1
        
        # ★★★ 修正天数指标计算（尾部）★★★
        days_total = df.loc[event_peak_date:end_date].shape[0]
        days_to_bottom = df.loc[event_peak_date:event_low_date].shape[0]
        # 尾部事件尚未结束，同理：回弹天数 = 总天数 - 回撤天数
        days_to_recovery = days_total - days_to_bottom
        # ***********************************

        events.append({
            "Start": event_peak_date.strftime("%Y-%m-%d"),
            "End": end_date.strftime("%Y-%m-%d"),
            "Bottom": event_low_date.strftime("%Y-%m-%d"),
            "PeakHigh": f"{event_peak_price:.2f}",
            "BottomLow": f"{event_low_price:.2f}",
            "MaxDD": f"{max_dd*100:.2f}%",
            "DaysTotal": days_total,
            "DaysToBottom": days_to_bottom,
            "DaysToRecovery": days_to_recovery,
        })

    print(f"\n{name} 回撤事件（修正天数计算逻辑）")
    if events:
        df_events = pd.DataFrame(events)
        # 调整列顺序，让天数指标更靠近 MaxDD
        cols = ["Start", "End", "Bottom", "DaysTotal", "DaysToBottom", "DaysToRecovery", "PeakHigh", "BottomLow", "MaxDD"]
        df_events = df_events[cols]
        print(df_events.to_string(index=False))
    else:
        print("(No events)")
    print("="*80)


# ============================================================
# 执行
# ============================================================
detect_events(spy, 0.05, "SPY -5%")
detect_events(qqq, 0.08, "QQQ -8%")
import math
from scipy.stats import norm
from datetime import datetime

# 期权参数（请替换为你要查的）
S = 231.44                      # 当前股价
K = 230                        # 行权价
T = (datetime(2025, 8, 22) - datetime.now()).days / 365  # 距离到期时间（以年为单位）
r = 0.05                       # 无风险利率（可调整）
iv = 0.39                      # 隐含波动率（用小数，如 0.39）
option_type = "put"           # 'call' or 'put'

# 计算 d1 和 d2
d1 = (math.log(S / K) + (r + 0.5 * iv**2) * T) / (iv * math.sqrt(T))
d2 = d1 - iv * math.sqrt(T)

# Delta
if option_type == "call":
    delta = norm.cdf(d1)
else:
    delta = norm.cdf(d1) - 1

# Gamma（call 和 put 相同）
gamma = norm.pdf(d1) / (S * iv * math.sqrt(T))

# Theta（每天）
if option_type == "call":
    theta = (-S * norm.pdf(d1) * iv / (2 * math.sqrt(T)) -
             r * K * math.exp(-r * T) * norm.cdf(d2)) / 365
else:
    theta = (-S * norm.pdf(d1) * iv / (2 * math.sqrt(T)) +
             r * K * math.exp(-r * T) * norm.cdf(-d2)) / 365

# 输出
print(f"Delta: {delta:.6f}")
print(f"Gamma: {gamma:.6f}")
print(f"Theta: {theta:.6f}")

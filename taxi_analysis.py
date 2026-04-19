"""
NYC Yellow Taxi — Demand, Fare, and Tip Mining
==============================================
CMPE 255 Data Mining class project.
Author: Anukrithi Myadala.

End-to-end pipeline on the real January 2024 Yellow Taxi Trip Records
from the NYC Taxi and Limousine Commission, joined with the TLC Taxi
Zone Lookup. Implements every major algorithm family covered in the
course from scratch in NumPy/pandas: five regression types (simple
linear, multi-variable linear, log-linear, Poisson, logistic),
decision-tree classification (CART/Gini), K-means clustering,
DBSCAN outlier detection, Apriori association-rule mining, and PCA.

All models are fit on a stratified-by-hour 400,000-trip sample to
keep runtime under five minutes while preserving the January 2024
demand profile. Every metric reported in the PPT and Word report is
reproduced by running this single script.
"""
import json
import os
from collections import Counter
from math import erfc

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

# --------------------------------------------------------------------
# 0. CONFIG
# --------------------------------------------------------------------
RANDOM_STATE = 42
np.random.seed(RANDOM_STATE)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CSV_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "yellow_tripdata_2024-01.csv"),
    os.path.expanduser("~/Downloads/yellow_tripdata_2024-01.csv"),
    os.path.expanduser("~/Downloads/mini2-claude/yellow_tripdata_2024-01.csv"),
]
ZONES_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "taxi_zone_lookup.csv"),
    os.path.expanduser("~/Downloads/taxi_zone_lookup.csv"),
]
CHART_DIR = os.path.join(SCRIPT_DIR, "charts")
os.makedirs(CHART_DIR, exist_ok=True)

SAMPLE_N = 400_000

# Styling
NAVY, TEAL, CREAM, AMBER, ROSE, SLATE = "#0F3B66", "#0D9488", "#F8FAFC", "#D97706", "#BE185D", "#334155"
sns.set_style("whitegrid")
plt.rcParams.update({
    "axes.edgecolor": "#CBD5E1", "axes.labelcolor": SLATE,
    "xtick.color": SLATE, "ytick.color": SLATE,
    "font.family": "DejaVu Sans", "savefig.dpi": 150,
    "savefig.bbox": "tight",
})

def savefig(name):
    p = os.path.join(CHART_DIR, name)
    plt.savefig(p, bbox_inches="tight")
    plt.close()
    return p

# --------------------------------------------------------------------
# 1. LOAD DATA
# --------------------------------------------------------------------
csv_path = next((p for p in CSV_CANDIDATES if os.path.exists(p)), None)
if csv_path is None:
    raise FileNotFoundError("yellow_tripdata_2024-01.csv not found.")
print(f"Loading {csv_path} ...")
raw = pd.read_csv(
    csv_path,
    parse_dates=["tpep_pickup_datetime", "tpep_dropoff_datetime"],
    low_memory=False,
)
zones_path = next((p for p in ZONES_CANDIDATES if os.path.exists(p)), None)
if zones_path is None:
    raise FileNotFoundError("taxi_zone_lookup.csv not found. Place it next to this script.")
zones = pd.read_csv(zones_path)
print(f"Raw rows: {len(raw):,}    columns: {list(raw.columns)}")
print(f"Zones: {len(zones)}")

# --------------------------------------------------------------------
# 2. CLEANING — documented audit trail
# --------------------------------------------------------------------
audit = []
def step(name, df):
    audit.append({"step": name, "rows": len(df)})
    print(f"  {name:<42s} rows = {len(df):>10,}")

step("0. raw", raw)
df = raw.copy()

df = df.dropna(subset=["tpep_pickup_datetime","tpep_dropoff_datetime",
                       "trip_distance","fare_amount","total_amount",
                       "PULocationID","DOLocationID"])
step("1. drop rows with core NULLs", df)

df = df[df["fare_amount"] >= 0]; step("2. fare_amount >= 0", df)
df = df[df["total_amount"] >= 0]; step("3. total_amount >= 0", df)
df = df[df["tip_amount"] >= 0];   step("4. tip_amount >= 0", df)
df = df[df["trip_distance"] >= 0]; step("5. trip_distance >= 0", df)
df = df[df["trip_distance"] <= 100]; step("6. trip_distance <= 100 mi", df)

df = df[df["tpep_pickup_datetime"] < df["tpep_dropoff_datetime"]]
step("7. pickup < dropoff", df)

start = pd.Timestamp("2024-01-01"); end = pd.Timestamp("2024-02-01")
df = df[(df["tpep_pickup_datetime"] >= start) & (df["tpep_pickup_datetime"] < end)]
step("8. pickup within Jan 2024", df)

df["duration_min"] = (df["tpep_dropoff_datetime"] - df["tpep_pickup_datetime"]).dt.total_seconds() / 60.0
df = df[(df["duration_min"] >= 1) & (df["duration_min"] <= 360)]
step("9. 1 min <= duration <= 6 hr", df)

df["passenger_count"] = df["passenger_count"].fillna(1)
df = df[(df["passenger_count"] >= 1) & (df["passenger_count"] <= 6)]
step("10. 1 <= passengers <= 6", df)

# Percentile-based outlier removal (1st–99th) — gentler than IQR which
# was removing all trips > 4.86 mi and cutting legitimate airport rides.
for col in ["fare_amount","trip_distance"]:
    lo, hi = df[col].quantile([0.01, 0.99])
    lo = max(lo, 0)
    df = df[(df[col] >= lo) & (df[col] <= hi)]
    step(f"11. pctl clip {col} [{lo:.2f}, {hi:.2f}]", df)

# --------------------------------------------------------------------
# 3. FEATURE ENGINEERING + ZONE JOIN
# --------------------------------------------------------------------
df["hour"] = df["tpep_pickup_datetime"].dt.hour
df["day_of_week"] = df["tpep_pickup_datetime"].dt.dayofweek
df["is_weekend"] = df["day_of_week"].isin([5,6]).astype(int)
df["is_rush_hour"] = df["hour"].isin([7,8,9,16,17,18,19]).astype(int)
df["speed_mph"] = df["trip_distance"] / (df["duration_min"] / 60.0)
df["speed_mph"] = df["speed_mph"].clip(0, 80)
df["has_tip"] = (df["tip_amount"] > 0).astype(int)
df["tip_pct"] = df["tip_amount"] / df["fare_amount"].replace(0, np.nan)
df["tip_pct"] = df["tip_pct"].fillna(0).clip(0, 1)

# Airport flag (JFK=132, LGA=138, EWR=1 — NYC TLC zone IDs)
AIRPORT_IDS = {132, 138, 1}
df["is_airport_pickup"] = df["PULocationID"].isin(AIRPORT_IDS).astype(int)
df["is_airport_dropoff"] = df["DOLocationID"].isin(AIRPORT_IDS).astype(int)

zmap_borough = dict(zip(zones["LocationID"], zones["Borough"]))
zmap_zone    = dict(zip(zones["LocationID"], zones["Zone"]))
df["pu_borough"] = df["PULocationID"].map(zmap_borough).fillna("Unknown")
df["do_borough"] = df["DOLocationID"].map(zmap_borough).fillna("Unknown")
df["pu_zone"]    = df["PULocationID"].map(zmap_zone).fillna("Unknown")
df["do_zone"]    = df["DOLocationID"].map(zmap_zone).fillna("Unknown")

print(f"After cleaning + features: rows = {len(df):,}")

# Keep a reference to the FULL cleaned dataset so EDA charts reflect reality.
# Population charts (hourly volume, DoW heatmap, borough volume) should NOT be
# drawn from the stratified sample — stratification makes hourly counts look flat.
df_full = df.copy()

# Stratified sample by hour for the modelling layer
if len(df) > SAMPLE_N:
    df = (df.groupby("hour", group_keys=False)
            .apply(lambda g: g.sample(
                n=min(len(g), int(SAMPLE_N/24)+1),
                random_state=RANDOM_STATE))
            .reset_index(drop=True))
    step(f"12. stratified sample (n={SAMPLE_N:,} by hour)", df)

audit_df = pd.DataFrame(audit)
audit_df.to_csv(os.path.join(SCRIPT_DIR, "cleaning_audit.csv"), index=False)

# --------------------------------------------------------------------
# 4. EDA CHARTS — drawn on FULL cleaned df so the real demand profile shows
# --------------------------------------------------------------------
# 4a. Trips by hour of day (population)
hourly = df_full.groupby("hour").size()
plt.figure(figsize=(7.8, 4.2))
bars = plt.bar(hourly.index, hourly.values/1000, color=NAVY, edgecolor="white")
plt.title(f"Trip Volume by Hour of Day ({len(df_full)/1e6:.2f} M cleaned trips)",
          fontsize=13, fontweight="bold", pad=12)
plt.xlabel("Hour of day"); plt.ylabel("Thousand trips")
plt.xticks(range(0,24))
# Highlight the peak hour
peak_hr = hourly.idxmax()
bars[peak_hr].set_color(AMBER)
plt.annotate(f"Peak: {peak_hr}:00\n{hourly.max()/1000:.0f}K trips",
             xy=(peak_hr, hourly.max()/1000),
             xytext=(peak_hr-5, hourly.max()/1000*0.95),
             fontsize=10, color=AMBER, fontweight="bold",
             arrowprops=dict(arrowstyle="->", color=AMBER))
plt.ylim(0, hourly.max()/1000 * 1.15)
savefig("01_hourly_volume.png")

# 4b. Heatmap day-of-week x hour (population)
heat = df_full.groupby(["day_of_week","hour"]).size().unstack(fill_value=0)
heat.index = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
plt.figure(figsize=(9, 3.8))
sns.heatmap(heat/1000, cmap="YlGnBu", cbar_kws={"label":"Thousand trips"},
            linewidths=0.2, linecolor="white")
plt.title(f"Demand Heatmap — Day of Week × Hour ({len(df_full)/1e6:.2f} M trips)",
          fontsize=13, fontweight="bold", pad=10)
plt.xlabel("Hour of day"); plt.ylabel("")
savefig("02_dow_hour_heatmap.png")

# 4c. Fare + distance distributions (full cleaned data, not sample)
fig, axes = plt.subplots(1, 2, figsize=(10, 3.8))
axes[0].hist(df_full["fare_amount"], bins=40, color=TEAL, edgecolor="white")
axes[0].set_title("Fare Amount Distribution", fontweight="bold", pad=10)
axes[0].set_xlabel("USD"); axes[0].set_ylabel("Trips")
axes[1].hist(df_full["trip_distance"], bins=40, color=AMBER, edgecolor="white")
axes[1].set_title("Trip Distance Distribution", fontweight="bold", pad=10)
axes[1].set_xlabel("Miles"); axes[1].set_ylabel("Trips")
plt.tight_layout()
savefig("03_fare_distance.png")

# 4d. Borough pickup volume — use full cleaned data so absolute counts are meaningful
bor = df_full["pu_borough"].value_counts().head(6)
plt.figure(figsize=(7.5, 4.0))
plt.bar(bor.index, bor.values/1000, color=[NAVY,TEAL,AMBER,ROSE,"#475569","#7c3aed"][:len(bor)])
plt.title("Pickup Volume by Borough (full cleaned month)",
          fontsize=13, fontweight="bold", pad=10)
plt.ylabel("Thousand trips")
ymax = bor.values.max()/1000 * 1.18
plt.ylim(0, ymax)
for i, v in enumerate(bor.values):
    plt.text(i, v/1000 + ymax*0.01, f"{v/1000:.0f}K", ha="center", fontsize=9, fontweight="bold")
savefig("04_borough_volume.png")

# 4e. Correlation heatmap — larger font, higher-resolution labels
num_cols = ["fare_amount","trip_distance","duration_min","passenger_count",
            "tip_amount","tip_pct","total_amount","speed_mph","hour","is_weekend","is_rush_hour"]
corr = df_full[num_cols].corr()
plt.figure(figsize=(8.2, 6.0))
sns.heatmap(corr, annot=True, fmt=".2f", cmap="RdBu_r", center=0, vmin=-1, vmax=1,
            cbar_kws={"label":"Pearson r"}, annot_kws={"size":11, "weight":"bold"},
            linewidths=0.3, linecolor="white")
plt.title("Numeric Feature Correlation (full cleaned data)",
          fontsize=13, fontweight="bold", pad=12)
plt.xticks(rotation=40, ha="right")
savefig("05_correlation.png")

# --------------------------------------------------------------------
# 5. FROM-SCRATCH ALGORITHMS
# --------------------------------------------------------------------

def train_test_split(X, y, test_size=0.2, stratify=None, seed=RANDOM_STATE):
    n = len(X)
    rng = np.random.default_rng(seed)
    if stratify is None:
        idx = rng.permutation(n)
        cut = int(n * (1 - test_size))
        return X[idx[:cut]], X[idx[cut:]], y[idx[:cut]], y[idx[cut:]]
    # stratified
    tr_idx, te_idx = [], []
    for cls in np.unique(stratify):
        where = np.where(stratify == cls)[0]
        rng.shuffle(where)
        cut = int(len(where)*(1-test_size))
        tr_idx.extend(where[:cut]); te_idx.extend(where[cut:])
    tr_idx = np.array(tr_idx); te_idx = np.array(te_idx)
    rng.shuffle(tr_idx); rng.shuffle(te_idx)
    return X[tr_idx], X[te_idx], y[tr_idx], y[te_idx]

def standardize(X_train, X_test):
    mu = X_train.mean(axis=0); sd = X_train.std(axis=0); sd[sd==0] = 1.0
    return (X_train - mu)/sd, (X_test - mu)/sd, mu, sd

def r2(y_true, y_pred):
    ss_res = np.sum((y_true - y_pred)**2)
    ss_tot = np.sum((y_true - y_true.mean())**2)
    return 1 - ss_res/ss_tot if ss_tot > 0 else 0.0

def rmse(y_true, y_pred): return float(np.sqrt(np.mean((y_true - y_pred)**2)))
def mae(y_true, y_pred):  return float(np.mean(np.abs(y_true - y_pred)))
def adj_r2(y_true, y_pred, p):
    n = len(y_true); rr = r2(y_true,y_pred)
    return 1 - (1-rr)*(n-1)/(n-p-1) if n > p+1 else rr

# ---- K-fold Cross-Validation (from scratch) ----
def kfold_indices(n, k=5, seed=RANDOM_STATE):
    """Generate k-fold train/test index splits."""
    rng = np.random.default_rng(seed)
    idx = rng.permutation(n)
    fold_size = n // k
    folds = []
    for i in range(k):
        te = idx[i * fold_size:(i + 1) * fold_size] if i < k - 1 else idx[i * fold_size:]
        tr = np.setdiff1d(idx, te)
        folds.append((tr, te))
    return folds

def cross_validate_regression(X, y, fit_fn, predict_fn, k=5, standardize_data=True):
    """K-fold CV for regression; returns mean and std of R^2 and RMSE."""
    folds = kfold_indices(len(X), k)
    r2s, rmses = [], []
    for tr, te in folds:
        Xtr, Xte, ytr, yte = X[tr], X[te], y[tr], y[te]
        if standardize_data:
            Xtr, Xte, _, _ = standardize(Xtr, Xte)
        w = fit_fn(Xtr, ytr)
        yp = predict_fn(w, Xte)
        r2s.append(r2(yte, yp))
        rmses.append(rmse(yte, yp))
    return {"r2_mean": np.mean(r2s), "r2_std": np.std(r2s),
            "rmse_mean": np.mean(rmses), "rmse_std": np.std(rmses)}

def cross_validate_classification(X, y, fit_fn, predict_proba_fn, k=5, threshold=0.5):
    """K-fold CV for classification; returns mean and std of AUC and F1."""
    folds = kfold_indices(len(X), k)
    aucs, f1s = [], []
    for tr, te in folds:
        Xtr, Xte, ytr, yte = X[tr], X[te], y[tr], y[te]
        Xtr_s, Xte_s, _, _ = standardize(Xtr, Xte)
        w = fit_fn(Xtr_s, ytr)
        prob = predict_proba_fn(w, Xte_s)
        aucs.append(roc_auc(yte, prob))
        pred = (prob >= threshold).astype(int)
        _, _, _, f1_val = prf(yte, pred)
        f1s.append(f1_val)
    return {"auc_mean": np.mean(aucs), "auc_std": np.std(aucs),
            "f1_mean": np.mean(f1s), "f1_std": np.std(f1s)}

# ---- Linear Regression (closed-form) ----
def fit_linear_regression(X, y):
    Xb = np.c_[np.ones(len(X)), X]
    w, *_ = np.linalg.lstsq(Xb, y, rcond=None)
    return w
def predict_linear(w, X): return np.c_[np.ones(len(X)), X] @ w

# ---- Logistic Regression (gradient descent, class-weighted) ----
def fit_logistic(X, y, class_weight="balanced", lr=0.05, n_iter=800, l2=1e-3):
    Xb = np.c_[np.ones(len(X)), X]
    n, p = Xb.shape
    w = np.zeros(p)
    if class_weight == "balanced":
        n_pos = max(1, int(y.sum())); n_neg = max(1, int(len(y)-y.sum()))
        wp = len(y)/(2*n_pos); wn = len(y)/(2*n_neg)
        sw = np.where(y==1, wp, wn)
    else:
        sw = np.ones(n)
    for _ in range(n_iter):
        z = np.clip(Xb @ w, -500, 500)
        p_hat = 1.0/(1.0+np.exp(-z))
        grad = (Xb.T @ ((p_hat - y)*sw))/n + l2*w
        w -= lr*grad
    return w
def predict_logistic_proba(w, X):
    z = np.clip(np.c_[np.ones(len(X)), X] @ w, -500, 500)
    return 1.0/(1.0+np.exp(-z))

# ---- Poisson Regression (IRLS, vectorized — no dense n×n matrix) ----
def fit_poisson(X, y, n_iter=25, eps=1e-6):
    Xb = np.c_[np.ones(len(X)), X]
    n, p = Xb.shape
    w = np.zeros(p); w[0] = np.log(max(y.mean(), 1e-3))
    for _ in range(n_iter):
        eta = np.clip(Xb @ w, -20, 20)
        mu = np.exp(eta) + eps
        z = eta + (y - mu) / mu
        # Vectorized weighted least squares: W is diagonal, so use element-wise multiply
        sqrt_w = np.sqrt(mu)
        Xw = Xb * sqrt_w[:, None]      # X * sqrt(W)
        zw = z * sqrt_w                 # z * sqrt(W)
        try:
            XtWX = Xw.T @ Xw + np.eye(p) * 1e-6
            XtWz = Xw.T @ zw
            w_new = np.linalg.solve(XtWX, XtWz)
        except np.linalg.LinAlgError:
            break
        if np.max(np.abs(w_new - w)) < 1e-6: break
        w = w_new
    return w
def predict_poisson(w, X):
    return np.exp(np.clip(np.c_[np.ones(len(X)), X] @ w, -20, 20))

# ---- Decision Tree (CART / Gini, binary classification) ----
class DecisionTree:
    def __init__(self, max_depth=6, min_samples_split=50, min_samples_leaf=25):
        self.max_depth = max_depth
        self.mss = min_samples_split
        self.msl = min_samples_leaf
    def _gini(self, y):
        if len(y)==0: return 0.0
        p = y.mean()
        return 1 - p*p - (1-p)*(1-p)
    def _best_split(self, X, y):
        n, d = X.shape
        best = None
        parent_g = self._gini(y)
        for j in range(d):
            col = X[:, j]
            vals = np.quantile(col, np.linspace(0.1, 0.9, 9))
            for thr in np.unique(vals):
                left = col <= thr; right = ~left
                if left.sum() < self.msl or right.sum() < self.msl: continue
                gl = self._gini(y[left]); gr = self._gini(y[right])
                g = (left.sum()*gl + right.sum()*gr)/n
                gain = parent_g - g
                if best is None or gain > best[0]:
                    best = (gain, j, thr)
        return best
    def _grow(self, X, y, depth):
        if depth >= self.max_depth or len(y) < self.mss or y.sum() in (0, len(y)):
            return {"leaf": True, "p": float(y.mean() if len(y) else 0.0), "n": int(len(y))}
        best = self._best_split(X, y)
        if best is None or best[0] <= 0:
            return {"leaf": True, "p": float(y.mean()), "n": int(len(y))}
        _, j, thr = best
        left = X[:, j] <= thr; right = ~left
        return {"leaf": False, "j": j, "thr": float(thr),
                "L": self._grow(X[left], y[left], depth+1),
                "R": self._grow(X[right], y[right], depth+1),
                "gain": float(best[0])}
    def fit(self, X, y):
        self.tree = self._grow(np.asarray(X, dtype=float), np.asarray(y, dtype=float), 0)
        return self
    def _pred_one(self, x, node):
        if node["leaf"]: return node["p"]
        return self._pred_one(x, node["L"] if x[node["j"]] <= node["thr"] else node["R"])
    def predict_proba(self, X):
        return np.array([self._pred_one(x, self.tree) for x in np.asarray(X, dtype=float)])
    def feature_importance(self, n_features):
        imp = np.zeros(n_features)
        def walk(node):
            if node["leaf"]: return
            imp[node["j"]] += node.get("gain", 0.0)
            walk(node["L"]); walk(node["R"])
        walk(self.tree)
        s = imp.sum()
        return imp/s if s > 0 else imp

# ---- K-means from scratch (k-means++ init, multiple restarts) ----
def _kmeanspp_init(X, k, rng):
    """K-means++ initialization for better starting centroids."""
    n = len(X)
    centers = [X[rng.integers(n)]]
    for _ in range(1, k):
        dists = np.min([np.sum((X - c)**2, axis=1) for c in centers], axis=0)
        probs = dists / dists.sum()
        centers.append(X[rng.choice(n, p=probs)])
    return np.array(centers)

def kmeans(X, k, n_iter=100, n_restarts=10, seed=RANDOM_STATE):
    """K-means with k-means++ init and multiple restarts (best inertia wins)."""
    best_lab, best_C, best_inertia = None, None, np.inf
    for r in range(n_restarts):
        rng = np.random.default_rng(seed + r)
        C = _kmeanspp_init(X, k, rng)
        for _ in range(n_iter):
            d = np.linalg.norm(X[:, None, :] - C[None, :, :], axis=2)
            lab = d.argmin(axis=1)
            new = np.array([X[lab==j].mean(axis=0) if (lab==j).any() else C[j] for j in range(k)])
            if np.allclose(new, C): break
            C = new
        inertia = float(((X - C[lab])**2).sum())
        if inertia < best_inertia:
            best_lab, best_C, best_inertia = lab.copy(), C.copy(), inertia
    return best_lab, best_C, best_inertia

def silhouette(X, labels, sample=2000, seed=RANDOM_STATE):
    rng = np.random.default_rng(seed)
    n = len(X); idx = rng.choice(n, size=min(sample, n), replace=False)
    Xs = X[idx]; ls = labels[idx]
    sils = []
    for i in range(len(Xs)):
        same = ls == ls[i]; same[i] = False
        if same.sum() == 0: continue
        a = np.linalg.norm(Xs[same] - Xs[i], axis=1).mean()
        b = np.inf
        for c in np.unique(ls):
            if c == ls[i]: continue
            mask = ls == c
            if mask.sum() == 0: continue
            dist = np.linalg.norm(Xs[mask] - Xs[i], axis=1).mean()
            b = min(b, dist)
        if b == np.inf: continue
        sils.append((b - a) / max(a, b))
    return float(np.mean(sils)) if sils else 0.0

# ---- DBSCAN (O(n^2) fine for n<1000) ----
def dbscan(X, eps, min_pts):
    n = len(X)
    labels = -np.ones(n, dtype=int)
    D = np.linalg.norm(X[:,None,:] - X[None,:,:], axis=2)
    neigh = [np.where(D[i] <= eps)[0] for i in range(n)]
    C = 0
    for i in range(n):
        if labels[i] != -1: continue
        if len(neigh[i]) < min_pts:
            labels[i] = -2  # noise (will remap to -1 at end)
            continue
        labels[i] = C
        seeds = list(neigh[i])
        while seeds:
            j = seeds.pop()
            if labels[j] == -2: labels[j] = C
            if labels[j] != -1: continue
            labels[j] = C
            if len(neigh[j]) >= min_pts:
                seeds.extend([k for k in neigh[j] if labels[k] == -1 or labels[k] == -2])
        C += 1
    labels[labels == -2] = -1
    return labels

# ---- Metrics for classification ----
def confusion_matrix_bin(y_true, y_pred):
    tp = int(((y_pred==1)&(y_true==1)).sum()); tn = int(((y_pred==0)&(y_true==0)).sum())
    fp = int(((y_pred==1)&(y_true==0)).sum()); fn = int(((y_pred==0)&(y_true==1)).sum())
    return np.array([[tn, fp],[fn, tp]]), tp, tn, fp, fn

def prf(y_true, y_pred):
    _, tp, tn, fp, fn = confusion_matrix_bin(y_true, y_pred)
    acc = (tp+tn)/max(1, tp+tn+fp+fn)
    prec = tp/max(1, tp+fp); rec = tp/max(1, tp+fn)
    f1 = 2*prec*rec/max(1e-9, prec+rec)
    return acc, prec, rec, f1

def roc_auc(y_true, score):
    order = np.argsort(-score)
    y = y_true[order]
    P = y.sum(); N = len(y) - P
    if P == 0 or N == 0: return 0.5
    tpr, fpr = [], []
    tp = fp = 0
    for yi in y:
        if yi == 1: tp += 1
        else: fp += 1
        tpr.append(tp/P); fpr.append(fp/N)
    # trapezoid
    fpr = np.array([0]+fpr); tpr = np.array([0]+tpr)
    return float(np.trapz(tpr, fpr))

# --------------------------------------------------------------------
# 6. REGRESSION STACK — 5 types
# --------------------------------------------------------------------
print("\n=== REGRESSION STACK ===")
reg_results = {}

# --- 6a. Simple linear: fare ~ distance ---
X1 = df[["trip_distance"]].values; y1 = df["fare_amount"].values.astype(float)
Xtr, Xte, ytr, yte = train_test_split(X1, y1)
w_s = fit_linear_regression(Xtr, ytr)
yp = predict_linear(w_s, Xte)
reg_results["Simple Linear (fare~distance)"] = {
    "r2": r2(yte, yp), "adj_r2": adj_r2(yte, yp, 1),
    "rmse": rmse(yte, yp), "mae": mae(yte, yp),
    "equation": f"fare = {w_s[0]:.3f} + {w_s[1]:.3f} * distance",
}

# --- 6b. Multi-variable linear: fare ~ many features ---
mv_feats = ["trip_distance","duration_min","passenger_count","hour",
            "is_weekend","is_rush_hour","is_airport_pickup","is_airport_dropoff"]
X2 = df[mv_feats].values.astype(float); y2 = df["fare_amount"].values.astype(float)
Xtr, Xte, ytr, yte = train_test_split(X2, y2)
Xtr_s, Xte_s, mu2, sd2 = standardize(Xtr, Xte)
# Save multi-var test set for residual plot later
Xte_mv, yte_mv = Xte.copy(), yte.copy()
w_m = fit_linear_regression(Xtr_s, ytr)
yp = predict_linear(w_m, Xte_s)
reg_results["Multi-Variable Linear (fare)"] = {
    "r2": r2(yte, yp), "adj_r2": adj_r2(yte, yp, len(mv_feats)),
    "rmse": rmse(yte, yp), "mae": mae(yte, yp),
    "top_coefs": dict(sorted(
        zip(mv_feats, w_m[1:]), key=lambda kv: -abs(kv[1])
    )[:5]),
}

# --- 6c. Log-linear: log(duration) ~ features ---
y3 = np.log(df["duration_min"].values + 1e-3)
X3 = df[["trip_distance","passenger_count","hour","is_weekend","is_rush_hour"]].values.astype(float)
Xtr, Xte, ytr, yte = train_test_split(X3, y3)
Xtr_s, Xte_s, _, _ = standardize(Xtr, Xte)
w_l = fit_linear_regression(Xtr_s, ytr)
yp_log = predict_linear(w_l, Xte_s)
yp_raw = np.exp(yp_log); yte_raw = np.exp(yte)
reg_results["Log-Linear (log duration)"] = {
    "r2_log_space": r2(yte, yp_log),
    "r2_raw_space": r2(yte_raw, yp_raw),
    "rmse_raw_min": rmse(yte_raw, yp_raw),
    "mae_raw_min": mae(yte_raw, yp_raw),
    "skewness_before_log": float(pd.Series(df["duration_min"]).skew()),
    "skewness_after_log": float(pd.Series(y3).skew()),
}

# --- 6d. Poisson: hourly trip counts per zone ---
counts = (df.groupby(["PULocationID","hour"]).size()
            .rename("trips").reset_index())
counts = counts.merge(zones[["LocationID","Borough"]],
                      left_on="PULocationID", right_on="LocationID", how="left")
counts["is_rush"] = counts["hour"].isin([7,8,9,16,17,18,19]).astype(int)
counts["is_manhattan"] = (counts["Borough"]=="Manhattan").astype(int)
counts["is_airport_zone"] = counts["PULocationID"].isin(AIRPORT_IDS).astype(int)
Xp = counts[["hour","is_rush","is_manhattan","is_airport_zone"]].values.astype(float)
yp_count = counts["trips"].values.astype(float)
Xtr, Xte, ytr, yte = train_test_split(Xp, yp_count)
Xtr_s, Xte_s, _, _ = standardize(Xtr, Xte)
w_p = fit_poisson(Xtr_s, ytr)
yp_pred = predict_poisson(w_p, Xte_s)
dev_null = 2*np.sum(yte*np.log((yte+1e-9)/ytr.mean()) - (yte - ytr.mean()))
dev_model = 2*np.sum(yte*np.log((yte+1e-9)/np.maximum(yp_pred,1e-9)) - (yte - yp_pred))
pseudo_r2 = 1 - dev_model/dev_null if dev_null != 0 else 0.0
dispersion_ratio = float(np.var(yte) / max(np.mean(yte), 1e-9))
reg_results["Poisson (hourly trip counts)"] = {
    "pseudo_r2_mcfadden": float(pseudo_r2),
    "rmse": rmse(yte, yp_pred),
    "mae": mae(yte, yp_pred),
    "mean_count": float(ytr.mean()),
    "dispersion": dispersion_ratio,
    "overdispersion_note": (
        f"Dispersion ratio = {dispersion_ratio:.1f} (Poisson assumes 1.0). "
        "Severe overdispersion — standard errors are underestimated. "
        "A Negative Binomial model would be more appropriate for these "
        "highly variable count data."
    ),
}

# --- 6e. Logistic: P(tip > 0) ---
# NOTE: payment_type is NOT used as a feature — in TLC data tips are only
# recorded for credit-card transactions, so including it would be data
# leakage (the feature is essentially a proxy for the label).
lr_feats = ["fare_amount","trip_distance","duration_min","passenger_count",
            "hour","is_weekend","is_rush_hour","is_airport_pickup",
            "speed_mph","is_airport_dropoff"]
lr_feats_full = lr_feats
Xl = df[lr_feats].values.astype(float)
yl = df["has_tip"].values.astype(int)
Xtr, Xte, ytr, yte = train_test_split(Xl, yl, stratify=yl)
Xtr_s, Xte_s, mu_lr, sd_lr = standardize(Xtr, Xte)
w_lr = fit_logistic(Xtr_s, ytr)
proba_lr = predict_logistic_proba(w_lr, Xte_s)
yte_lr = yte.copy()
auc_lr = roc_auc(yte_lr, proba_lr)
pred_lr = (proba_lr >= 0.5).astype(int)
acc, pr, rc, f1 = prf(yte_lr, pred_lr)
reg_results["Logistic (P(tip>0))"] = {
    "accuracy": acc, "precision": pr, "recall": rc, "f1": f1, "auc": auc_lr,
}

# Save feature importances
coef_df = pd.DataFrame({"feature": lr_feats_full, "coef": w_lr[1:]}).sort_values("coef", key=abs, ascending=False)
coef_df.to_csv(os.path.join(SCRIPT_DIR, "top_features.csv"), index=False)

# --- 5-fold cross-validation for key models ---
print("\n=== 5-FOLD CROSS-VALIDATION ===")
cv_simple = cross_validate_regression(X1, y1, fit_linear_regression, predict_linear, standardize_data=False)
cv_multi  = cross_validate_regression(X2, y2, fit_linear_regression, predict_linear, standardize_data=True)
cv_logistic = cross_validate_classification(Xl, yl, fit_logistic, predict_logistic_proba)
print(f"  Simple Linear  5-fold R²: {cv_simple['r2_mean']:.4f} ± {cv_simple['r2_std']:.4f}")
print(f"  Multi-Var Lin  5-fold R²: {cv_multi['r2_mean']:.4f} ± {cv_multi['r2_std']:.4f}")
print(f"  Logistic tip   5-fold AUC: {cv_logistic['auc_mean']:.4f} ± {cv_logistic['auc_std']:.4f}")
reg_results["Simple Linear (fare~distance)"]["cv_r2"] = f"{cv_simple['r2_mean']:.4f} ± {cv_simple['r2_std']:.4f}"
reg_results["Multi-Variable Linear (fare)"]["cv_r2"] = f"{cv_multi['r2_mean']:.4f} ± {cv_multi['r2_std']:.4f}"
reg_results["Logistic (P(tip>0))"]["cv_auc"] = f"{cv_logistic['auc_mean']:.4f} ± {cv_logistic['auc_std']:.4f}"

# Save regression comparison
reg_rows = []
for k, v in reg_results.items():
    row = {"model": k}
    row.update({kk: (round(vv,4) if isinstance(vv,(int,float)) else str(vv)[:120])
                for kk, vv in v.items() if not isinstance(vv, dict)})
    reg_rows.append(row)
pd.DataFrame(reg_rows).to_csv(os.path.join(SCRIPT_DIR, "model_comparison.csv"), index=False)

# --------------------------------------------------------------------
# 7. DECISION TREE — peak-hour vs off-peak classifier
# --------------------------------------------------------------------
print("\n=== DECISION TREE (peak-hour classifier) ===")
dt_feats = ["trip_distance","duration_min","fare_amount","passenger_count",
            "is_weekend","is_airport_pickup","is_airport_dropoff","speed_mph"]
Xd = df[dt_feats].values.astype(float)
yd = df["is_rush_hour"].values.astype(int)
Xtr, Xte, ytr, yte = train_test_split(Xd, yd, stratify=yd)
dt = DecisionTree(max_depth=6).fit(Xtr, ytr)
proba_dt = dt.predict_proba(Xte)
yte_dt = yte.copy()
# Use class-prevalence threshold so the model actually predicts the positive class
prev = float(ytr.mean())
pred_dt = (proba_dt >= prev).astype(int)
acc_dt, pr_dt, rc_dt, f1_dt = prf(yte_dt, pred_dt)
auc_dt = roc_auc(yte_dt, proba_dt)
dt_cm, *_ = confusion_matrix_bin(yte_dt, pred_dt)
dt_imp = dt.feature_importance(len(dt_feats))
print(f"DT: acc={acc_dt:.3f} prec={pr_dt:.3f} rec={rc_dt:.3f} f1={f1_dt:.3f} auc={auc_dt:.3f}")

# --------------------------------------------------------------------
# 8. K-MEANS — zone archetypes
# --------------------------------------------------------------------
print("\n=== K-MEANS zone archetypes ===")
zone_agg = (df.groupby("PULocationID")
              .agg(trips=("fare_amount","size"),
                   avg_fare=("fare_amount","mean"),
                   avg_distance=("trip_distance","mean"),
                   avg_duration=("duration_min","mean"),
                   morning_ratio=("hour", lambda s: (s.between(6,11)).mean()),
                   evening_ratio=("hour", lambda s: (s.between(17,22)).mean()),
                   airport_share=("is_airport_pickup","mean"))
              .reset_index())
zone_agg = zone_agg[zone_agg["trips"] >= 50].copy()
feats_z = ["trips","avg_fare","avg_distance","avg_duration",
           "morning_ratio","evening_ratio","airport_share"]
Xz = zone_agg[feats_z].values.astype(float)
mu_z = Xz.mean(axis=0); sd_z = Xz.std(axis=0); sd_z[sd_z==0]=1.0
Xz_s = (Xz - mu_z)/sd_z

# elbow scan + automatic elbow detection (max curvature / kneedle heuristic)
elbow = {}
for k in range(2, 10):
    _, _, inertia = kmeans(Xz_s, k)
    elbow[k] = inertia

# Pick k at the point of maximum curvature (second derivative)
ks_arr = np.array(sorted(elbow.keys()))
inertias_arr = np.array([elbow[k] for k in ks_arr])
# Normalize to [0,1] for fair curvature comparison
kn = (ks_arr - ks_arr[0]) / (ks_arr[-1] - ks_arr[0])
yn = (inertias_arr - inertias_arr[-1]) / (inertias_arr[0] - inertias_arr[-1])
# Curvature = distance from line connecting first and last point
distances = np.abs(yn - (1 - kn))  # distance to y = 1-x diagonal
best_k = int(ks_arr[np.argmax(distances)])
print(f"Elbow auto-detected at k = {best_k}")
labels_z, centers_z, _ = kmeans(Xz_s, best_k)
sil = silhouette(Xz_s, labels_z, sample=min(800, len(Xz_s)))
zone_agg["cluster"] = labels_z
zone_agg = zone_agg.merge(zones[["LocationID","Borough","Zone"]],
                          left_on="PULocationID", right_on="LocationID", how="left")

# Build interpretable profiles
profiles = (zone_agg.groupby("cluster")[feats_z].mean().round(2)
            .assign(size=zone_agg.groupby("cluster").size()))
# Force distinct labels by ranking on each archetype dimension and assigning
# the strongest cluster to each label exactly once.
labels_unique = ["Airport Feeder", "Morning Commute", "Nightlife / Evening",
                 "Dense High-Volume Core", "Long-Haul Premium",
                 "Low-Volume Periphery", "Residential / Mixed"]
score_keys = {
    "Airport Feeder":        lambda p: p["airport_share"],
    "Morning Commute":       lambda p: p["morning_ratio"] - p["evening_ratio"],
    "Nightlife / Evening":   lambda p: p["evening_ratio"] - p["morning_ratio"],
    "Dense High-Volume Core":lambda p: p["trips"],
    "Long-Haul Premium":     lambda p: p["avg_fare"] * p["avg_distance"],
    "Low-Volume Periphery": lambda p: -p["trips"],
}
remaining = list(profiles.index)
final_label = {}
for lab in labels_unique[:-1]:
    if not remaining: break
    if lab not in score_keys: continue
    scores = {c: score_keys[lab](profiles.loc[c]) for c in remaining}
    best = max(scores, key=scores.get)
    final_label[best] = lab
    remaining.remove(best)
for c in remaining:
    final_label[c] = "Residential / Mixed"
profiles["label"] = [final_label[c] for c in profiles.index]
zone_agg["cluster_label"] = zone_agg["cluster"].map(profiles["label"].to_dict())
profiles.to_csv(os.path.join(SCRIPT_DIR, "cluster_profiles.csv"))
print(profiles[["label","size"]].to_string())
print(f"Silhouette: {sil:.3f}")

# --------------------------------------------------------------------
# 9. DBSCAN — outlier hotspot zones
# --------------------------------------------------------------------
print("\n=== DBSCAN zone outliers ===")
db_labels = dbscan(Xz_s, eps=1.2, min_pts=4)
n_clusters_db = len(set(db_labels)) - (1 if -1 in db_labels else 0)
n_noise = int((db_labels == -1).sum())
zone_agg["dbscan"] = db_labels
print(f"DBSCAN clusters: {n_clusters_db}, outliers: {n_noise}")
# Pull the actual outlier zone names so the slide narrative is truthful
outlier_zones = (zone_agg[zone_agg["dbscan"] == -1]
                 .sort_values("trips", ascending=False)
                 [["Zone", "Borough", "trips", "avg_fare", "airport_share"]]
                 .head(8))
outlier_zones.to_csv(os.path.join(SCRIPT_DIR, "dbscan_outliers.csv"), index=False)
print("Top DBSCAN outlier zones:")
print(outlier_zones.to_string(index=False))

# --------------------------------------------------------------------
# 10. APRIORI — pickup→dropoff zone-level association rules
# --------------------------------------------------------------------
print("\n=== APRIORI (zone-level rules) ===")

# Each trip is a transaction with items:
#   PU_zone:<name>, DO_zone:<name>, PU_bor:<bor>, DO_bor:<bor>,
#   rush:<0/1>, weekend:<0/1>
# This gives a richer item space for real Apriori candidate generation.

def apriori_from_scratch(transactions, min_sup=0.005, min_conf=0.3):
    """True Apriori: level-wise candidate generation with downward-closure pruning."""
    n = len(transactions)
    min_count = int(n * min_sup)

    # --- Pass 1: frequent 1-itemsets ---
    item_counts = Counter()
    for t in transactions:
        for item in t:
            item_counts[item] += 1
    freq1 = {frozenset([item]): cnt for item, cnt in item_counts.items() if cnt >= min_count}
    all_freq = dict(freq1)
    prev_freq = freq1

    # --- Level-wise: generate k-itemsets from (k-1)-itemsets ---
    k = 2
    while prev_freq and k <= 4:
        # Candidate generation: join step
        prev_items = list(prev_freq.keys())
        candidates = set()
        for i in range(len(prev_items)):
            for j in range(i + 1, len(prev_items)):
                c = prev_items[i] | prev_items[j]
                if len(c) == k:
                    # Pruning: all (k-1)-subsets must be frequent (downward closure)
                    subsets_freq = all(
                        (c - frozenset([item])) in prev_freq
                        for item in c
                    )
                    if subsets_freq:
                        candidates.add(c)

        if not candidates:
            break

        # Count support for candidates
        cand_counts = Counter()
        for t in transactions:
            t_set = frozenset(t)
            for c in candidates:
                if c.issubset(t_set):
                    cand_counts[c] += 1

        prev_freq = {c: cnt for c, cnt in cand_counts.items() if cnt >= min_count}
        all_freq.update(prev_freq)
        k += 1

    # --- Rule generation ---
    rules = []
    for itemset, sup_count in all_freq.items():
        if len(itemset) < 2:
            continue
        sup = sup_count / n
        for item in itemset:
            antecedent = itemset - frozenset([item])
            consequent = frozenset([item])
            ant_count = all_freq.get(antecedent, 0)
            if ant_count == 0:
                continue
            conf = sup_count / ant_count
            if conf < min_conf:
                continue
            cons_sup = all_freq.get(consequent, 1) / n
            lift_val = conf / cons_sup if cons_sup > 0 else 0
            rules.append({
                "antecedent": ", ".join(sorted(antecedent)),
                "consequent": ", ".join(sorted(consequent)),
                "support": round(sup, 4),
                "confidence": round(conf, 3),
                "lift": round(lift_val, 2),
                "count": sup_count,
            })
    return rules, all_freq

# Build transaction list with zone-level items (vectorized — no iterrows)
top_pu_set = set(df["pu_zone"].value_counts().head(40).index)
top_do_set = set(df["do_zone"].value_counts().head(40).index)

pu_z = np.where(df["pu_zone"].isin(top_pu_set), "PU=" + df["pu_zone"].astype(str), "")
do_z = np.where(df["do_zone"].isin(top_do_set), "DO=" + df["do_zone"].astype(str), "")
pu_b = ("PU_bor=" + df["pu_borough"].astype(str)).values
do_b = ("DO_bor=" + df["do_borough"].astype(str)).values
rush = np.where(df["is_rush_hour"].values == 1, "rush=yes", "")
wknd = np.where(df["is_weekend"].values == 1, "weekend=yes", "")

# Stack columns and build sets, filtering empty strings
all_items = np.column_stack([pu_z, do_z, pu_b, do_b, rush, wknd])
transactions = [set(row[row != ""]) for row in all_items]
del all_items

rules, freq_itemsets = apriori_from_scratch(transactions, min_sup=0.005, min_conf=0.3)
print(f"Frequent itemsets found: {len(freq_itemsets)}")
if rules:
    rules_df = pd.DataFrame(rules).sort_values("lift", ascending=False)
    rules_df = rules_df[~rules_df["antecedent"].str.contains("Unknown") &
                        ~rules_df["consequent"].str.contains("Unknown")].reset_index(drop=True)
    rules_df.to_csv(os.path.join(SCRIPT_DIR, "association_rules.csv"), index=False)
    print(f"Association rules (conf >= 0.3): {len(rules_df)}")
    print(rules_df.head(15).to_string(index=False))
else:
    rules_df = pd.DataFrame(columns=["antecedent","consequent","support","confidence","lift","count"])
    rules_df.to_csv(os.path.join(SCRIPT_DIR, "association_rules.csv"), index=False)
    print("  No rules met thresholds — try lowering min_sup/min_conf")

# --------------------------------------------------------------------
# 11. PCA variance analysis
# --------------------------------------------------------------------
print("\n=== PCA variance ===")
pca_feats = ["fare_amount","trip_distance","duration_min","passenger_count",
             "tip_amount","total_amount","speed_mph","hour"]
Xp_all = df[pca_feats].values.astype(float)
mu_p = Xp_all.mean(axis=0); sd_p = Xp_all.std(axis=0); sd_p[sd_p==0]=1
Xp_n = (Xp_all - mu_p)/sd_p
cov = np.cov(Xp_n.T)
eigvals, eigvecs = np.linalg.eigh(cov)
idx = np.argsort(-eigvals)
eigvals = eigvals[idx]; eigvecs = eigvecs[:, idx]
var_ratio = eigvals/eigvals.sum()
cum_var = np.cumsum(var_ratio)
n_components_90 = int(np.argmax(cum_var >= 0.90) + 1)
print(f"Components for 90% variance: {n_components_90} / {len(pca_feats)}")

# --------------------------------------------------------------------
# 12. DIAGNOSTIC CHARTS for presentation
# --------------------------------------------------------------------
# 12a. Regression results bar chart
r2_bars = {
    "Simple Linear\n(fare~dist)": reg_results["Simple Linear (fare~distance)"]["r2"],
    "Multi-Var Linear\n(fare)": reg_results["Multi-Variable Linear (fare)"]["r2"],
    "Log-Linear\n(duration)": reg_results["Log-Linear (log duration)"]["r2_log_space"],
    "Poisson\n(counts, pseudo-R²)": reg_results["Poisson (hourly trip counts)"]["pseudo_r2_mcfadden"],
    "Logistic\n(tip, AUC)": reg_results["Logistic (P(tip>0))"]["auc"],
}
plt.figure(figsize=(8.5, 4.2))
bars = list(r2_bars.keys()); vals = list(r2_bars.values())
cols = [NAVY, TEAL, AMBER, ROSE, "#7c3aed"]
b = plt.bar(bars, vals, color=cols)
ymax = max(vals)*1.2
plt.ylim(0, max(ymax, 1.0))
plt.title("Regression & Classification Scores Across Five Model Families",
          fontsize=12, fontweight="bold", pad=12)
plt.ylabel("R² (or AUC for logistic)")
for rect, v in zip(b, vals):
    plt.text(rect.get_x()+rect.get_width()/2, v+0.02, f"{v:.3f}",
             ha="center", fontsize=10, fontweight="bold")
savefig("06_regression_scores.png")

# 12b. Decision-tree confusion matrix + feature importance
fig, ax = plt.subplots(1, 2, figsize=(10, 4.2))
sns.heatmap(dt_cm, annot=True, fmt="d", cmap="Blues", cbar=False,
            xticklabels=["Off-peak","Peak"], yticklabels=["Off-peak","Peak"],
            ax=ax[0], annot_kws={"size":12, "weight":"bold"})
ax[0].set_title(f"Decision Tree Confusion Matrix (AUC={auc_dt:.3f})",
                fontweight="bold", pad=10)
ax[0].set_xlabel("Predicted"); ax[0].set_ylabel("Actual")
order = np.argsort(dt_imp)
ax[1].barh(np.array(dt_feats)[order], dt_imp[order], color=TEAL)
ax[1].set_title("Feature Importance", fontweight="bold", pad=10)
ax[1].set_xlabel("Gini-gain share")
plt.tight_layout()
savefig("07_dt_cm_importance.png")

# 12c. K-means scatter (PC1 vs PC2 of zone features colored by cluster)
# Compute PCA on zone feature matrix
cov_z = np.cov(Xz_s.T)
ez, vz = np.linalg.eigh(cov_z); iz = np.argsort(-ez)
Vz = vz[:, iz[:2]]
coords = Xz_s @ Vz
plt.figure(figsize=(8, 5))
label_map = profiles["label"].to_dict()
colors_k = plt.get_cmap("tab10")(np.arange(best_k))
for c in range(best_k):
    m = labels_z == c
    plt.scatter(coords[m,0], coords[m,1], s=46, color=colors_k[c],
                label=label_map.get(c, f"Cluster {c}"), edgecolor="white", linewidth=0.6)
plt.title("K-Means Zone Archetypes (projected onto 2 principal components)",
          fontsize=12, fontweight="bold", pad=12)
plt.xlabel("PC 1"); plt.ylabel("PC 2")
plt.legend(loc="best", fontsize=9, frameon=True)
savefig("08_kmeans_zones.png")

# 12d. DBSCAN outliers
plt.figure(figsize=(8, 5))
noise = db_labels == -1
plt.scatter(coords[~noise,0], coords[~noise,1], s=40, color=TEAL,
            label=f"In cluster (n={(~noise).sum()})", alpha=0.8)
plt.scatter(coords[noise,0], coords[noise,1], s=70, color=AMBER,
            marker="X", label=f"Outlier zone (n={noise.sum()})", edgecolor="black", linewidth=0.6)
plt.title("DBSCAN — Outlier Zones Flagged", fontsize=12, fontweight="bold", pad=12)
plt.xlabel("PC 1"); plt.ylabel("PC 2")
plt.legend(fontsize=10)
savefig("09_dbscan.png")

# 12e. Apriori top rules (zone-level)
if len(rules_df) > 0:
    top_rules = rules_df.head(15)
    plt.figure(figsize=(9.5, 5.6))
    labels_r = [f"{r['antecedent']} → {r['consequent']}" for _, r in top_rules.iterrows()]
    plt.barh(labels_r[::-1], top_rules["lift"].values[::-1], color=NAVY)
    plt.title("Top 15 Association Rules by Lift (Apriori, zone-level)",
              fontsize=12, fontweight="bold", pad=12)
    plt.xlabel("Lift")
    plt.tight_layout()
    savefig("10_apriori_rules.png")
else:
    print("  WARNING: No association rules found — try lowering min_sup or min_conf")

# 12f. Lift & cumulative gains for logistic tip classifier
# Sort by predicted probability desc, plot cumulative gains
order = np.argsort(-proba_lr)
y_sorted = yte_lr[order]
cum_pos = np.cumsum(y_sorted)
total_pos = max(y_sorted.sum(), 1)
n_ = len(y_sorted)
pct_pop = np.arange(1, n_+1)/n_
pct_pos = cum_pos/total_pos
baseline = pct_pop
lift = (pct_pos / np.maximum(pct_pop, 1e-9))
# Compute actual decile headline metrics for slide narrative
def at(p):
    i = int(n_*p) - 1
    return float(pct_pos[i]*100), float(lift[i])
cap20, lift20 = at(0.20); cap10, lift10 = at(0.10); cap30, lift30 = at(0.30)
print(f"Lift: top 10% captures {cap10:.1f}% (lift {lift10:.2f}x), "
      f"top 20% captures {cap20:.1f}% (lift {lift20:.2f}x)")

fig, ax = plt.subplots(1, 2, figsize=(11, 4.4))
ax[0].plot(pct_pop*100, pct_pos*100, color=NAVY, linewidth=2.5, label="Logistic model")
ax[0].plot(pct_pop*100, baseline*100, color="#94A3B8", linewidth=1.4, linestyle="--", label="Random baseline")
ax[0].fill_between(pct_pop*100, pct_pos*100, baseline*100, color=NAVY, alpha=0.10)
ax[0].axvline(20, color=AMBER, linestyle=":", linewidth=1.2)
ax[0].annotate(f"Top 20% → {cap20:.0f}% captured", xy=(20, cap20),
               xytext=(32, cap20+18), fontsize=10, color=AMBER, fontweight="bold",
               arrowprops=dict(arrowstyle="->", color=AMBER))
ax[0].set_title("Cumulative Gains Curve — Tip Classifier",
                fontweight="bold", pad=10)
ax[0].set_xlabel("% Population Targeted"); ax[0].set_ylabel("% Positive Captured")
ax[0].legend(loc="lower right"); ax[0].set_xlim(0,100); ax[0].set_ylim(0,100)
ax[0].grid(True, alpha=0.3)

ax[1].plot(pct_pop*100, lift, color=AMBER, linewidth=2.5)
ax[1].axhline(1.0, color="#94A3B8", linestyle="--", linewidth=1)
ax[1].axvline(20, color=NAVY, linestyle=":", linewidth=1.2)
ax[1].annotate(f"Lift at 20% = {lift20:.2f}x", xy=(20, lift20),
               xytext=(30, lift20+0.02), fontsize=10, color=NAVY, fontweight="bold",
               arrowprops=dict(arrowstyle="->", color=NAVY))
ax[1].set_title("Lift Chart — Tip Classifier", fontweight="bold", pad=10)
ax[1].set_xlabel("% Population Targeted"); ax[1].set_ylabel("Lift")
ax[1].set_xlim(0,100)
ax[1].grid(True, alpha=0.3)
plt.tight_layout()
savefig("11_lift_gains.png")

# 12f-extra. ROC curves for logistic and decision tree
def roc_points(y_true, score):
    order = np.argsort(-score); y = y_true[order]
    P = y.sum(); N = len(y)-P
    tp = fp = 0; tprs, fprs = [0], [0]
    for yi in y:
        if yi == 1: tp += 1
        else: fp += 1
        tprs.append(tp/P if P else 0); fprs.append(fp/N if N else 0)
    return np.array(fprs), np.array(tprs)
fpr_l, tpr_l = roc_points(yte_lr, proba_lr)
fpr_d, tpr_d = roc_points(yte_dt, proba_dt)
plt.figure(figsize=(6.5, 5.0))
plt.plot(fpr_l, tpr_l, color=NAVY, linewidth=2.4, label=f"Logistic (AUC={auc_lr:.3f})")
plt.plot(fpr_d, tpr_d, color=AMBER, linewidth=2.4, label=f"Decision Tree (AUC={auc_dt:.3f})")
plt.plot([0,1],[0,1], color="#94A3B8", linestyle="--", linewidth=1, label="Random")
plt.fill_between(fpr_l, tpr_l, alpha=0.08, color=NAVY)
plt.title("ROC Curves — Classification Models", fontsize=13, fontweight="bold", pad=10)
plt.xlabel("False Positive Rate"); plt.ylabel("True Positive Rate")
plt.legend(loc="lower right"); plt.xlim(0,1); plt.ylim(0,1)
plt.grid(True, alpha=0.3)
savefig("14_roc_curves.png")

# 12f-extra. K-means elbow + silhouette plot
plt.figure(figsize=(8.2, 3.2))
ks = sorted(elbow.keys())
inertias = [elbow[k] for k in ks]
plt.plot(ks, inertias, marker="o", color=NAVY, linewidth=2, markersize=8)
plt.axvline(best_k, color=AMBER, linestyle="--", linewidth=1.5,
            label=f"Selected k = {best_k}")
plt.title("K-Means Elbow Method — Inertia vs k",
          fontsize=12, fontweight="bold", pad=8)
plt.xlabel("Number of clusters (k)"); plt.ylabel("Inertia")
plt.xticks(ks); plt.legend(fontsize=9); plt.grid(True, alpha=0.3)
savefig("15_kmeans_elbow.png")

# 12g. PCA cumulative variance
plt.figure(figsize=(7.5, 3.8))
plt.plot(range(1, len(var_ratio)+1), cum_var*100, marker="o", color=NAVY)
plt.axhline(90, color=AMBER, linestyle="--", label="90% threshold")
plt.title("PCA — Cumulative Explained Variance", fontweight="bold", pad=10)
plt.xlabel("# Components"); plt.ylabel("% variance captured")
plt.legend()
savefig("12_pca_variance.png")

# 12h. Residual plot for multi-var linear (uses the SAME test split as model evaluation)
Xte_s_resid = (Xte_mv - mu2) / sd2
yp_mv_test = predict_linear(w_m, Xte_s_resid)[:1500]
resid = yte_mv[:1500] - yp_mv_test
plt.figure(figsize=(8, 3.8))
plt.scatter(yp_mv_test, resid, s=8, alpha=0.35, color=TEAL)
plt.axhline(0, color=SLATE, linewidth=1)
plt.title("Residuals — Multi-Variable Linear Regression",
          fontweight="bold", pad=10)
plt.xlabel("Predicted fare ($)"); plt.ylabel("Residual ($)")
savefig("13_residuals.png")

# --------------------------------------------------------------------
# 12i. HYPOTHESIS TESTING — statistical significance of key findings
# --------------------------------------------------------------------
print("\n=== HYPOTHESIS TESTING ===")
hypothesis_results = {}

# H1: Airport trips have higher fares than non-airport trips (two-sample t-test)
airport_fares = df[df["is_airport_pickup"] == 1]["fare_amount"].values
non_airport_fares = df[df["is_airport_pickup"] == 0]["fare_amount"].values
# Welch's t-test (from scratch — unequal variances)
def welch_t_test(x1, x2):
    n1, n2 = len(x1), len(x2)
    m1, m2 = x1.mean(), x2.mean()
    v1, v2 = x1.var(ddof=1), x2.var(ddof=1)
    se = np.sqrt(v1/n1 + v2/n2)
    t_stat = (m1 - m2) / se if se > 0 else 0
    # Welch-Satterthwaite degrees of freedom
    num = (v1/n1 + v2/n2)**2
    den = (v1/n1)**2/(n1-1) + (v2/n2)**2/(n2-1)
    dof = num / den if den > 0 else 1
    # p-value approximation using normal (valid for large n)
    p_value = erfc(abs(t_stat) / np.sqrt(2))
    return t_stat, p_value, dof

t1, p1, dof1 = welch_t_test(airport_fares, non_airport_fares)
hypothesis_results["H1: Airport fares > non-airport"] = {
    "t_statistic": round(t1, 3),
    "p_value": f"{p1:.2e}",
    "mean_airport": round(float(airport_fares.mean()), 2),
    "mean_non_airport": round(float(non_airport_fares.mean()), 2),
    "significant": p1 < 0.05,
}
print(f"  H1 airport fares: t={t1:.3f}, p={p1:.2e} → {'SIGNIFICANT' if p1<0.05 else 'not significant'}")

# H2: Rush-hour trips are longer in duration than off-peak (t-test)
rush_dur = df[df["is_rush_hour"] == 1]["duration_min"].values
off_dur  = df[df["is_rush_hour"] == 0]["duration_min"].values
t2, p2, _ = welch_t_test(rush_dur, off_dur)
hypothesis_results["H2: Rush-hour duration > off-peak"] = {
    "t_statistic": round(t2, 3),
    "p_value": f"{p2:.2e}",
    "mean_rush": round(float(rush_dur.mean()), 2),
    "mean_off_peak": round(float(off_dur.mean()), 2),
    "significant": p2 < 0.05,
}
print(f"  H2 rush duration: t={t2:.3f}, p={p2:.2e} → {'SIGNIFICANT' if p2<0.05 else 'not significant'}")

# H3: Weekend trips have different tip rates than weekday (chi-square test)
def chi_square_2x2(a, b, c, d):
    """2x2 contingency: [[a,b],[c,d]] chi-square."""
    n = a + b + c + d
    e1 = (a+b)*(a+c)/n; e2 = (a+b)*(b+d)/n
    e3 = (c+d)*(a+c)/n; e4 = (c+d)*(b+d)/n
    chi2 = sum((o-e)**2/e for o, e in [(a,e1),(b,e2),(c,e3),(d,e4)] if e > 0)
    p = erfc(np.sqrt(chi2/2))  # approximation for 1 dof
    return chi2, p

wkend_tip = int(df[(df["is_weekend"]==1) & (df["has_tip"]==1)].shape[0])
wkend_notip = int(df[(df["is_weekend"]==1) & (df["has_tip"]==0)].shape[0])
wkday_tip = int(df[(df["is_weekend"]==0) & (df["has_tip"]==1)].shape[0])
wkday_notip = int(df[(df["is_weekend"]==0) & (df["has_tip"]==0)].shape[0])
chi2, p3 = chi_square_2x2(wkend_tip, wkend_notip, wkday_tip, wkday_notip)
hypothesis_results["H3: Weekend tip rate != weekday"] = {
    "chi_square": round(chi2, 3),
    "p_value": f"{p3:.2e}",
    "weekend_tip_rate": round(wkend_tip / max(1, wkend_tip + wkend_notip), 4),
    "weekday_tip_rate": round(wkday_tip / max(1, wkday_tip + wkday_notip), 4),
    "significant": p3 < 0.05,
}
print(f"  H3 weekend tipping: chi²={chi2:.3f}, p={p3:.2e} → {'SIGNIFICANT' if p3<0.05 else 'not significant'}")

# --------------------------------------------------------------------
# 13. EXPORT SAMPLE FOR DASHBOARD
# --------------------------------------------------------------------
dash_cols = ["hour","day_of_week","is_weekend","is_rush_hour",
             "pu_borough","do_borough","pu_zone","do_zone","PULocationID",
             "trip_distance","duration_min","fare_amount","tip_amount","tip_pct",
             "total_amount","passenger_count","has_tip","speed_mph",
             "is_airport_pickup","is_airport_dropoff"]
dash = df[dash_cols].round(3)  # full modelling sample, 388K rows
# merge cluster label
dash = dash.merge(zone_agg[["PULocationID","cluster","cluster_label"]],
                  on="PULocationID", how="left")
dash["cluster_label"] = dash["cluster_label"].fillna("Residential / Mixed")
dash.to_csv(os.path.join(SCRIPT_DIR, "dashboard_sample.csv"), index=False)

# --------------------------------------------------------------------
# 14. WRITE A SUMMARY JSON
# --------------------------------------------------------------------
summary = {
    "rows_raw": int(len(raw)),
    "rows_clean": int(len(df_full)),
    "sample_n": int(len(df)),
    "regression_results": {k: {kk: (float(vv) if isinstance(vv,(int,float,np.floating)) else str(vv))
                               for kk, vv in v.items() if not isinstance(vv, dict)}
                            for k, v in reg_results.items()},
    "decision_tree": {"accuracy": acc_dt, "precision": pr_dt, "recall": rc_dt,
                      "f1": f1_dt, "auc": auc_dt},
    "kmeans": {"k": best_k, "silhouette": sil,
               "profiles": profiles[["label","size"]].reset_index().to_dict(orient="records")},
    "dbscan": {"clusters": n_clusters_db, "outliers": n_noise},
    "apriori_top3": rules_df.head(3).to_dict(orient="records"),
    "pca_components_for_90pct": int(n_components_90),
    "lift_metrics": {
        "top10_capture_pct": round(cap10,1), "top10_lift": round(lift10,3),
        "top20_capture_pct": round(cap20,1), "top20_lift": round(lift20,3),
        "top30_capture_pct": round(cap30,1), "top30_lift": round(lift30,3),
    },
    "dbscan_outlier_zones": outlier_zones.head(5).to_dict(orient="records"),
    "dt_feature_importance": {f: float(dt_imp[i]) for i, f in enumerate(dt_feats)},
    "kmeans_elbow": {str(k): float(v) for k, v in elbow.items()},
    "hypothesis_tests": hypothesis_results,
    "cross_validation": {
        "simple_linear_cv_r2": reg_results["Simple Linear (fare~distance)"].get("cv_r2", ""),
        "multi_var_cv_r2": reg_results["Multi-Variable Linear (fare)"].get("cv_r2", ""),
        "logistic_cv_auc": reg_results["Logistic (P(tip>0))"].get("cv_auc", ""),
    },
}
with open(os.path.join(SCRIPT_DIR, "summary.json"), "w") as f:
    json.dump(summary, f, indent=2, default=str)

print("\nDone. Artifacts written to:", SCRIPT_DIR)

import streamlit as st

st.set_page_config(page_title="溶液調製計算機 Pro", layout="centered")

# --- 入力セクション ---
# 単位の倍率定義
units_mol = {"mol/L": 1.0, "mmol/L": 1e-3, "μmol/L": 1e-6}
units_vol = {"L": 1.0, "mL": 1e-3, "μL": 1e-6}

col1, col2, col3 = st.columns(3)

with col1:
    z = st.number_input("分子量 z (g/mol)", value=1.0, format="%.2f")

with col2:
    x_val = st.number_input("目標濃度 x", value=0.0, format="%.4f")
    x_unit = st.selectbox("濃度の単位", list(units_mol.keys()))

with col3:
    y_val = st.number_input("目標体積 y", value=0.0, format="%.4f")
    y_unit = st.selectbox("体積の単位", list(units_vol.keys()))

# --- 計算ロジック ---
# すべて基本単位（mol/L, L）に変換して計算
real_x = x_val * units_mol[x_unit]
real_y = y_val * units_vol[y_unit]
res_g = real_x * real_y * z

st.divider()

# --- 結果表示セクション ---
st.subheader("📊 計算結果一覧")

# 同時に複数の単位で結果を表示
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("必要質量 (g)", f"{res_g:.4f} g")
with c2:
    st.metric("必要質量 (mg)", f"{res_g * 1e3:.2f} mg")
with c3:
    st.metric("必要質量 (μg)", f"{res_g * 1e6:.1f} μg")

# 入力値の再確認（ミス防止用）
st.caption(f"現在の設定: {x_val} {x_unit} × {y_val} {y_unit} (分子量: {z})")
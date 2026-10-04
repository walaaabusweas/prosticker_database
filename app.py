"""
app.py  —  واجهة إدخال بيانات المنتجات (Streamlit + SQLite)

تشغيل:   streamlit run app.py
"""
import io
import json
import os

import pandas as pd
import streamlit as st

import core
from core import FieldSpec

st.set_page_config(page_title="نظام إدخال بيانات المنتجات", page_icon="🖨️", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
html, body, .stApp, .stMarkdown, p, label, h1, h2, h3, input, textarea, button { font-family: 'Tajawal', sans-serif; }
.stApp, [data-testid="stSidebar"] { direction: rtl; }
.stApp h1, .stApp h2, .stApp h3, .stApp p, .stApp label { text-align: right; }
div[data-baseweb="select"], div[data-baseweb="select"] * { direction: rtl; }
[data-testid="stDataFrame"], [data-testid="stJson"], code, pre { direction: ltr; text-align: left; }
input, textarea { direction: auto; }
</style>
""",
    unsafe_allow_html=True,
)

ss = st.session_state
PAGE_ENTRY, PAGE_LIST, PAGE_EXPORT = "➕ إدخال / تعديل منتج", "📋 المنتجات المُدخلة", "⬇️ التصدير والنسخ الاحتياطي"
PAGES = [PAGE_ENTRY, PAGE_LIST, PAGE_EXPORT]
PRICE_TYPE = {"fixed_price": "fixed", "custom_formula": "fixed", "quote_price": "quote"}


# ----------------------------------------------------------------------------
# تهيئة
# ----------------------------------------------------------------------------
def check_password() -> bool:
    pw = os.environ.get("APP_PASSWORD")
    if not pw:
        try:
            pw = st.secrets.get("APP_PASSWORD")
        except Exception:
            pw = None
    if not pw or ss.get("authed"):
        return True
    entered = st.text_input("🔒 كلمة المرور", type="password")
    if entered and entered == pw:
        ss["authed"] = True
        st.rerun()
    elif entered:
        st.error("كلمة المرور غير صحيحة")
    return False


@st.cache_resource
def get_schema(mtime: float):
    return core.load_schema()


if not check_password():
    st.stop()

core.init_db()
schema = get_schema(core.SCHEMA_FILE.stat().st_mtime)

ss.setdefault("ver", 0)          # يتغير بعد كل حفظ/تحميل لتصفير الحقول
ss.setdefault("edit_id", None)
ss.setdefault("page", PAGE_ENTRY)
ss.setdefault("flash", None)

PRICE_SPECS = {
    "price_method": FieldSpec("price_method", "طريقة حساب السعر", "select",
                              [v for v, _ in schema.pricing], True, "core", True),
    "price_currency": FieldSpec("price_currency", "العملة", "select", core.CURRENCIES, False, "core", True),
    "price": FieldSpec("price", "السعر", "number", [], False, "core", True),
    "setup_fee": FieldSpec("setup_fee", "رسوم التجهيز (اختياري)", "number", [], False, "core", True),
    "pricing_formula": FieldSpec("pricing_formula", "معادلة التسعير", "text", [], False, "core", True),
}


# ----------------------------------------------------------------------------
# مفاتيح الحقول + التسميات
# ----------------------------------------------------------------------------
def fkey(spec: FieldSpec, ptype: str, subtype: str) -> str:
    base = f"f{ss.ver}|"
    return base + (spec.name if spec.stable else f"{ptype}|{subtype}|{spec.name}")


def dkey(spec, ptype, subtype) -> str:
    return fkey(spec, ptype, subtype) + "|detail"


def sub_key(ptype: str) -> str:
    return f"f{ss.ver}|sel_subtype|{ptype}"


def formatter(spec: FieldSpec):
    labels = schema.labels

    def _f(v):
        if v in ("", None):
            return "— اختر —"
        if spec.kind == "bool":
            return labels["bool"].get(v, v)
        return labels.get(spec.name, {}).get(v) or labels["generic"].get(v, v)

    return _f


def type_label(t: str) -> str:
    return "— اختر —" if not t else f"{t} — {schema.types[t]['ar']}"


# ----------------------------------------------------------------------------
# رسم حقل واحد (يرجع القيمة المخزنة)
# ----------------------------------------------------------------------------
def render_field(spec: FieldSpec, ptype: str, subtype: str, allow_blank: bool = True):
    key = fkey(spec, ptype, subtype)
    label = spec.label + (" *" if spec.required else "")
    hint = spec.name
    fmt = formatter(spec)

    if spec.kind == "text":
        state = st.text_input(label, key=key, help=hint)
    elif spec.kind == "textarea":
        state = st.text_area(label, key=key, height=90, help=hint)
    elif spec.kind == "number":
        state = st.number_input(label, value=None, min_value=0.0, step=1.0, format="%g", key=key, help=hint)
    elif spec.kind == "bool":
        state = st.selectbox(label, ["", "Yes", "No"], format_func=fmt, key=key, help=hint)
    elif spec.kind == "select":
        opts = ([""] if allow_blank else []) + spec.options
        state = st.selectbox(label, opts, format_func=fmt, key=key, help=hint)
    else:  # multi
        state = st.multiselect(label, spec.options, format_func=fmt, key=key, help=hint,
                               placeholder="اختر واحدًا أو أكثر")

    detail = ""
    if spec.kind in ("select", "multi") and spec.other_tokens:
        picked = state if isinstance(state, list) else [state]
        if any(t in picked for t in spec.other_tokens):
            detail = st.text_input(f"↳ تفاصيل: {spec.label}", key=dkey(spec, ptype, subtype))
    return core.state_to_value(spec, state, detail)


def render_grid(specs, ptype, subtype, values, ncols=3, skip=()):
    cols = st.columns(ncols)
    i = 0
    for sp in specs:
        if sp.name in skip:
            values[sp.name] = None
            continue
        with cols[i % ncols]:
            values[sp.name] = render_field(sp, ptype, subtype)
        i += 1


# ----------------------------------------------------------------------------
# Callbacks
# ----------------------------------------------------------------------------
def _set_state(sp, ptype, subtype, stored):
    state, detail = core.value_to_state(sp, stored)
    ss[fkey(sp, ptype, subtype)] = state
    if sp.other_tokens:
        ss[dkey(sp, ptype, subtype)] = detail


def load_record(pid: int, duplicate: bool):
    res = core.get_product(pid)
    if not res:
        return
    values, faqs = res
    ss.ver += 1
    ss.edit_id = None if duplicate else pid
    ptype, subtype = values.get("product_type", ""), values.get("subtype", "")
    if ptype not in schema.types:
        return
    ss["sel_type"] = ptype
    if subtype in schema.subtypes.get(ptype, []):
        ss[sub_key(ptype)] = subtype
    specs = core.build_specs(schema, ptype, subtype)
    for sp in core.all_specs(specs):
        _set_state(sp, ptype, subtype, values.get(sp.name))
    for name, sp in PRICE_SPECS.items():
        _set_state(sp, ptype, subtype, values.get(name))
    ss[f"faq_init|{ss.ver}"] = faqs
    ss.page = PAGE_ENTRY
    ss.flash = ("info", "تم تحميل المنتج للتعديل" if not duplicate else "تم نسخ المنتج — عدّل ما يلزم واحفظه كمنتج جديد")


def cancel_edit():
    ss.ver += 1
    ss.edit_id = None


def delete_cb(pid: int):
    core.delete_product(pid)
    ss.flash = ("success", f"تم حذف المنتج #{pid}")


def prefill_catalog():
    i = ss.get(f"cat_pick|{ss.ver}")
    if i in (None, ""):
        return
    e = schema.catalog[i]
    ptype, subtype = e["Final Type"], e.get("Subtype") or ""
    if ptype not in schema.types:
        return
    ss["sel_type"] = ptype
    if subtype in schema.subtypes.get(ptype, []):
        ss[sub_key(ptype)] = subtype
    ss[f"f{ss.ver}|product_name"] = e["Product / Entry Name"]
    core_specs = core.build_specs(schema, ptype, subtype)["core"]
    for col, name, multi in (("Printing Method", "printing_method", False),
                             ("Customization Method", "customization_method", True),
                             ("Occasion", "occasion", True)):
        val, sp = e.get(col), core_specs.get(name)
        if val and sp and val in sp.options:
            ss[fkey(sp, ptype, subtype)] = [val] if multi else val


# ----------------------------------------------------------------------------
# صفحة الإدخال
# ----------------------------------------------------------------------------
def page_entry():
    if ss.edit_id:
        c1, c2 = st.columns([4, 1])
        c1.warning(f"✏️ وضع التعديل — المنتج رقم {ss.edit_id}")
        c2.button("إلغاء التعديل", on_click=cancel_edit)

    with st.expander("⚡ ابدأ من منتج موجود في الكتالوج (اختياري)"):
        st.selectbox(
            "اختر منتجًا من الكتالوج ليتعبأ النوع والاسم تلقائيًا",
            [""] + list(range(len(schema.catalog))),
            format_func=lambda i: "— اختر —" if i == "" else
            f"{schema.catalog[i]['Product / Entry Name']}  ({schema.catalog[i]['Final Type']}/{schema.catalog[i]['Subtype']})",
            key=f"cat_pick|{ss.ver}", on_change=prefill_catalog,
        )

    c1, c2 = st.columns(2)
    ptype = c1.selectbox("نوع المنتج (Product Type) *", [""] + list(schema.types), format_func=type_label, key="sel_type")
    if not ptype:
        st.info("اختر نوع المنتج للبدء.")
        return
    subtype = c2.selectbox("النوع الفرعي (Subtype) *", [""] + schema.subtypes.get(ptype, []),
                           format_func=lambda v: "— اختر —" if not v else v, key=sub_key(ptype))
    if not subtype:
        st.info("اختر النوع الفرعي لتظهر خصائصه.")
        return

    specs = core.build_specs(schema, ptype, subtype)
    core_specs = specs["core"]
    values = {"product_type": ptype, "subtype": subtype}

    # 1) الأساسيات
    with st.container(border=True):
        st.subheader("1) المعلومات الأساسية")
        values["product_name"] = render_field(core_specs["product_name"], ptype, subtype)
        values["description"] = render_field(core_specs["description"], ptype, subtype)

    # 2) المواصفات العامة
    spec_names = ["material", "size", "custom_design", "design_file", "printing_method",
                  "customization_method", "printing_sides", "finish", "occasion"]
    spec_list = [core_specs[n] for n in spec_names if n in core_specs]
    cd = core_specs.get("custom_design")
    design_hidden = ("design_file",) if (cd and ss.get(fkey(cd, ptype, subtype)) != "Yes") else ()
    with st.container(border=True):
        st.subheader("2) المواصفات")
        render_grid(spec_list, ptype, subtype, values, skip=design_hidden)

    # 3) الكمية والتنفيذ
    qty_list = [core_specs[n] for n in ("quantity", "quantity_unit", "minimum_quantity", "production_time")
                if n in core_specs]
    with st.container(border=True):
        st.subheader("3) الكمية والتنفيذ")
        render_grid(qty_list, ptype, subtype, values, ncols=4)

    # 4) خصائص خاصة بالنوع / النوع الفرعي
    if specs["type"]:
        with st.expander(f"4) خصائص النوع {ptype}", expanded=True):
            render_grid(specs["type"], ptype, subtype, values)
    if specs["subtype"]:
        with st.expander(f"خصائص خاصة بـ {subtype}", expanded=True):
            render_grid(specs["subtype"], ptype, subtype, values)

    # 5) التسعير
    with st.container(border=True):
        st.subheader("5) التسعير")
        c1, c2 = st.columns(2)
        with c1:
            pm = render_field(PRICE_SPECS["price_method"], ptype, subtype)
        with c2:
            values["price_currency"] = render_field(PRICE_SPECS["price_currency"], ptype, subtype, allow_blank=False)
        values["price_method"] = pm
        values["pricing_type"] = PRICE_TYPE.get(pm)
        formula_error = None
        if pm == "fixed_price":
            values["price"] = render_field(PRICE_SPECS["price"], ptype, subtype)
            values["fixed_price"] = values["price"]
        if pm == "custom_formula":
            values["pricing_formula"] = render_field(PRICE_SPECS["pricing_formula"], ptype, subtype)
            variables, formula_error = core.analyze_formula(values["pricing_formula"] or "")
            if formula_error:
                st.error(formula_error)
            elif variables:
                st.caption("المتغيرات المكتشفة: " + ", ".join(variables))
                values["formula_variables"] = ", ".join(variables)
            st.caption("مثال: (width * height * quantity * material_rate) + setup_fee")
        if pm in ("fixed_price", "custom_formula"):
            values["setup_fee"] = render_field(PRICE_SPECS["setup_fee"], ptype, subtype)
        values["quote_required"] = 1 if pm == "quote_price" else (0 if pm else None)

    # 6) معلومات للبوت
    with st.container(border=True):
        st.subheader("6) معلومات تساعد البوت")
        c1, c2 = st.columns(2)
        with c1:
            values["use_cases"] = render_field(core_specs["use_cases"], ptype, subtype)
        with c2:
            values["target_businesses"] = render_field(core_specs["target_businesses"], ptype, subtype)
        st.markdown("**الأسئلة الشائعة (FAQ)** — أضف صفًا لكل سؤال وجوابه")
        init = ss.get(f"faq_init|{ss.ver}") or []
        faq_df = pd.DataFrame(
            {"السؤال": pd.Series([q for q, _ in init], dtype="str"),
             "الجواب": pd.Series([a for _, a in init], dtype="str")}
        )
        edited = st.data_editor(faq_df, num_rows="dynamic", key=f"faq|{ss.ver}")
        values["notes"] = render_field(core_specs["notes"], ptype, subtype)

    # حفظ
    if st.button("💾 حفظ المنتج", type="primary"):
        errors = []
        for sp in core.all_specs(specs) + [PRICE_SPECS["price_method"]]:
            if sp.required and core.is_empty(values.get(sp.name)):
                errors.append(sp.label)
        if pm == "fixed_price" and core.is_empty(values.get("price")):
            errors.append("السعر")
        if pm == "custom_formula" and (not values.get("pricing_formula") or formula_error):
            errors.append("معادلة التسعير (صحيحة)")
        if errors:
            st.error("أكمل الحقول المطلوبة: " + "، ".join(dict.fromkeys(errors)))
            return
        faqs = []
        for _, r in edited.iterrows():
            q, a = str(r.get("السؤال") or "").strip(), str(r.get("الجواب") or "").strip()
            if q and q.lower() != "nan":
                faqs.append((q, "" if a.lower() == "nan" else a))
        pid = core.save_product(values, faqs, ss.edit_id)
        ss.flash = ("success", f"✅ تم حفظ المنتج «{values['product_name']}» (رقم {pid})")
        ss.ver += 1
        ss.edit_id = None
        st.rerun()


# ----------------------------------------------------------------------------
# صفحة القائمة
# ----------------------------------------------------------------------------
def page_list():
    c1, c2 = st.columns([2, 1])
    search = c1.text_input("🔎 بحث بالاسم أو الوصف")
    tf = c2.selectbox("فلتر النوع", [""] + list(schema.types), format_func=type_label)
    rows = core.list_products(search, tf)
    st.caption(f"عدد المنتجات: {len(rows)}")
    if not rows:
        st.info("لا توجد منتجات بعد.")
        return
    df = pd.DataFrame(rows).rename(columns={
        "id": "#", "product_name": "الاسم", "product_type": "النوع", "subtype": "الفرعي",
        "printing_method": "الطباعة", "price_method": "التسعير", "price": "السعر",
        "quantity": "الكمية", "quantity_unit": "الوحدة", "updated_at": "آخر تعديل"})
    st.dataframe(df, hide_index=True)

    names = {r["id"]: r["product_name"] for r in rows}
    pick = st.selectbox("اختر منتجًا للإجراءات", list(names), format_func=lambda i: f"#{i} — {names[i]}")
    values, faqs = core.get_product(pick)
    c = st.columns(4)
    c[0].button("✏️ تعديل", on_click=load_record, args=(pick, False))
    c[1].button("📄 نسخ كمنتج جديد", on_click=load_record, args=(pick, True))
    ok = c[2].checkbox("تأكيد الحذف", key=f"del_ok|{pick}")
    c[3].button("🗑️ حذف", disabled=not ok, on_click=delete_cb, args=(pick,))
    with st.expander("التفاصيل الكاملة", expanded=True):
        st.json({**values, "faq": [{"question": q, "answer": a} for q, a in faqs]})


# ----------------------------------------------------------------------------
# صفحة التصدير
# ----------------------------------------------------------------------------
def page_export():
    df = core.export_dataframe()
    st.caption(f"عدد المنتجات: {len(df)}")
    if df.empty:
        st.info("لا توجد بيانات للتصدير بعد.")
        return
    st.dataframe(df, hide_index=True)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    c = st.columns(4)
    c[0].download_button("⬇️ Excel", buf.getvalue(), "products.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c[1].download_button("⬇️ CSV", df.to_csv(index=False).encode("utf-8-sig"), "products.csv", "text/csv")
    c[2].download_button("⬇️ JSON (للبوت)", json.dumps(core.export_records(), ensure_ascii=False, indent=2).encode("utf-8"),
                         "products.json", "application/json")
    with open(core.DB_FILE, "rb") as f:
        c[3].download_button("⬇️ قاعدة البيانات (.db)", f.read(), "catalog.db", "application/octet-stream")
    st.caption("أرسل ملف الـ .db أو الـ JSON للمطوّر عند الانتهاء من التعبئة.")


# ----------------------------------------------------------------------------
# التشغيل
# ----------------------------------------------------------------------------
def main():
    st.title("🖨️ نظام إدخال بيانات المنتجات")
    counts = core.counts_by_type()
    with st.sidebar:
        st.radio("القائمة", PAGES, key="page")
        st.metric("إجمالي المنتجات", sum(counts.values()))
        st.markdown("**التقدم حسب النوع**")
        for t in schema.types:
            n = counts.get(t, 0)
            st.write(f"{'✅' if n else '⚪'} {t} — {schema.types[t]['ar']}: **{n}**")

    if ss.flash:
        kind, msg = ss.flash
        getattr(st, kind)(msg)
        ss.flash = None

    {PAGE_ENTRY: page_entry, PAGE_LIST: page_list, PAGE_EXPORT: page_export}[ss.page]()


main()

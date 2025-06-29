import pandas as pd
import openai
import time
import re

# OpenAI Setup
# Make sure to replace with your actual API key or load from env/config

# Load the Excel sheet
df = pd.read_excel(".xlsx")

# Normalize column names: remove leading/trailing spaces, make lowercase
df.columns = df.columns.str.strip().str.lower()

# Ensure content columns are strings to avoid type issues
df["meta_description"] = df["meta_description"].astype(str)
df["product_description"] = df["product_description"].astype(str)

# Cleaning Utilities

def clean_meta_description(meta_text):
    """
    Clean the meta description by removing non-Arabic characters.
    """
    return re.sub(r"[^\u0600-\u06FF\s]", "", meta_text).strip()

def clean_product_description(description_text):
    """
    Clean the product description by removing unwanted symbols like asterisks and quotes.
    """
    return description_text.replace("**", "").replace("*", "").replace('"', "").strip()

# Prompt Templates

def prompt_full_description(product_name):
    """
    Create a prompt to generate the full marketing content for the product.
    """
    return f"""
    أنت كاتب محتوى تسويقي محترف. مهمتك هي كتابة وصف منتج احترافي باللغة العربية للمنتج التالي:

    اسم المنتج: {product_name}

    اكتب النص بصيغة رسمية وجذابة، مع مراعاة التنسيق التالي تمامًا:

    وصف المنتج:
    اكتب فقرتين تسويقيتين تشرحان المنتج ومكوناته وفوائده العامة، باستخدام لغة واضحة وجاذبة، بدون رموز مثل " أو *.

    المميزات الرئيسية:
    اكتب 4 إلى 6 نقاط موجزة تسلط الضوء على أهم مميزات المنتج وفوائده، مع الحفاظ على تنسيق النقاط وبدون رموز خاصة.  

    طريقة الاستخدام:
    اشرح للمستخدم كيفية استخدام المنتج بطريقة واضحة، خطوة بخطوة إذا لزم الأمر.
    """

def prompt_meta_description(product_name):
    """
    Create a prompt to generate a short, SEO-friendly Arabic meta description.
    """
    return f"""
    اكتب وصفاً تسويقياً موجزاً باللغة العربية للمنتج التالي:

    اسم المنتج: {product_name}
    اكتب وصفاً تسويقياً موجزاً باللغة العربية للمنتج التالي، لا يتجاوز 150 حرفاً. يجب أن يكون الوصف جذاباً، يعبر بوضوح عن فائدة المنتج، ومناسباً لمحركات البحث. لا تستخدم رموزاً أو علامات ترقيم
    """

# Content Generation Logic

def generate_content(product_name):
    """
    Generates the product description and meta description using two separate GPT-4o prompts.
    """

    # Generate full product description
    response_desc = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt_full_description(product_name)}],
        temperature=0.7,
        max_tokens=700
    )
    full_description = response_desc.choices[0].message.content.strip()
    description_cleaned = clean_product_description(full_description)

    # Generate SEO-friendly meta description
    response_meta = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt_meta_description(product_name)}],
        temperature=0.7,
        max_tokens=200
    )
    meta_description = response_meta.choices[0].message.content.strip()
    meta_cleaned = clean_meta_description(meta_description)

    # Preview in console
    print("\n Full Description:\n", full_description)
    print("\n Meta Description:\n", meta_description)

    return description_cleaned, meta_cleaned

# Process Each Product

for index, row in df.iterrows():
    product_name = row["product_name"]
    print(f"\n Processing row {index}: {product_name}")

    # Generate new content
    description, meta = generate_content(product_name)

    # See raw meta result before cleaning
    print("Meta (raw):", repr(meta))

    # Save results into the DataFrame
    df.at[index, "product_description"] = description
    df.at[index, "meta_description"] = meta

    # Save progress after each row to avoid losing work
    df.to_excel("- filled.xlsx", index=False)
    print(f" Row saved {index} to Excel.")

    # Small delayto the API
    time.sleep(1)

print("\n All Done!")
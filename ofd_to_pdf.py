# coding: utf-8
import streamlit as st
import base64
from easyofd.ofd import OFD
import easyofd.draw.draw_pdf as draw_pdf
import os
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
import io
import sys
import subprocess
# -------------------------

original_draw_signature = draw_pdf.DrawPDF.draw_signature
def safe_draw_signature(self, *args, **kwargs):
    try:
        return original_draw_signature(self, *args, **kwargs)
    except Exception as e:
        print(f"警告: 签章解析跳过 (Error: {e})")
        return None
draw_pdf.DrawPDF.draw_signature = safe_draw_signature

original_call = draw_pdf.DrawPDF.__call__
def safe_call(self, *args, **kwargs):
    try:
        return original_call(self, *args, **kwargs)
    except Exception as e:
        print(f"警告: draw_pdf.__call__ 跳过错误 (Error: {e})")
        return None
draw_pdf.DrawPDF.__call__ = safe_call

original_draw_annotation = draw_pdf.DrawPDF.draw_annotation
def safe_draw_annotation(self, c, annotation_info, images, page_size):
    try:
        safe_annotations = []
        for annotation in annotation_info:
            img_obj = annotation.get("ImgageObject")
            if img_obj is None:
                print("警告: 跳过空 ImgObject 注释")
                continue
            boundary = img_obj.get("Boundary", "")
            if boundary:
                boundary_parts = boundary.split(" ")
            else:
                boundary_parts = []
            annotation["ImgageObject"]["BoundaryParts"] = boundary_parts
            safe_annotations.append(annotation)
        return original_draw_annotation(self, c, safe_annotations, images, page_size)
    except Exception as e:
        print(f"警告: draw_annotation 异常跳过 ({e})")
        return None
draw_pdf.DrawPDF.draw_annotation = safe_draw_annotation

# -------------------------
# Streamlit 页面
# -------------------------
st.set_page_config(page_title="OFD 批量转 PDF 并打包", page_icon="📑")

st.title("🚀 OFD 批量转 PDF 版本（ZIP 下载）")
st.markdown("""
支持批量处理和多线程处理，原始文件和生成 PDF ，并打包成 ZIP 下载。
""")

# -------------------------
# 本地保存文件夹
# -------------------------
def create_output_folder():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join("ofd_output", timestamp)
    os.makedirs(folder, exist_ok=True)
    return folder

# -------------------------
# OFD 转 PDF 核心函数
# -------------------------
def convert_and_save(file, output_folder):
    try:
        file_bytes = file.read()
        filename_base = os.path.splitext(file.name)[0]

        # 保存原始 OFD
        ofd_path = os.path.join(output_folder, file.name)
        with open(ofd_path, "wb") as f:
            f.write(file_bytes)

        # 转 PDF
        ofd_obj = OFD()
        ofd_obj.read(base64.b64encode(file_bytes).decode("utf-8"))
        pdf_bytes = ofd_obj.to_pdf()
        ofd_obj.del_data()

        # 保存 PDF
        pdf_path = os.path.join(output_folder, f"{filename_base}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        return file.name, pdf_path
    except Exception as e:
        return file.name, None
# -------------------------
# 批量文件上传
# -------------------------
uploaded_files = st.file_uploader(
    "上传多个 OFD 文件", type=["ofd"], accept_multiple_files=True
)

if uploaded_files:
    st.write(f"共上传 {len(uploaded_files)} 个文件")

    if st.button("开始批量转换"):
        output_folder = create_output_folder()
        #st.info(f"文件将保存到本地: {output_folder}")

        results = []
        with st.spinner("正在转换文件..."):
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(convert_and_save, file, output_folder) for file in uploaded_files]
                for future in futures:
                    results.append(future.result())

        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            for filename, pdf_path in results:
                if pdf_path:
                    zipf.write(pdf_path, os.path.basename(pdf_path))
        zip_buffer.seek(0)

        st.success("✅ 批量转换完成！")
        st.download_button(
            label="💾 下载全部 PDF (ZIP)",
            data=zip_buffer,
            file_name="ofd_pdfs.zip",
            mime="application/zip"
        )

        for filename, pdf_path in results:
            if pdf_path:
                st.write(f"{filename} → 转换成功")
            else:
                st.write(f"{filename} → 转换失败")

st.divider()
st.info("提示：如果 PDF 中文字显示为方框，说明系统缺少中文字体（如宋体）。"
    "请在服务器上安装常用中文字体以正常显示。")
# -------------------------
# 自动启动 Streamlit，绑定 /ofdtopdf
# -------------------------
if __name__ == "__main__":
    cmd = f"streamlit run {sys.argv[0]} " \
          f"--server.address 192.168.7.36 " \
          f"--server.port 8501 " \
          f"--server.baseUrlPath /ofdtopdf " \
          f"--server.headless true"
    print(f"启动 Streamlit 应用: {cmd}")
    subprocess.run(cmd, shell=True)
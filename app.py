import streamlit as st
import tempfile
import subprocess
import os
import shutil
from io import BytesIO

# Cấu hình trang
st.set_page_config(
    page_title="Markdown to Word Converter",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tùy chỉnh với theme teal
st.markdown("""
<style>
    /* Import fonts */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    
    /* Root variables */
    :root {
        --primary-teal: #008080;
        --light-teal: #20b2aa;
        --dark-teal: #006666;
        --accent-teal: #40e0d0;
        --teal-bg: #f0fdfd;
        --teal-light: #e6ffff;
    }
    
    /* Main container */
    .main {
        padding: 2rem;
        font-family: 'Poppins', sans-serif;
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, var(--primary-teal), var(--light-teal));
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0, 128, 128, 0.2);
    }
    
    .main-header h1 {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        text-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    
    .main-header p {
        color: rgba(255, 255, 255, 0.9);
        font-size: 1.1rem;
        margin: 0.5rem 0 0 0;
        font-weight: 300;
    }
    
    /* Card styling */
    .card {
        background: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 128, 128, 0.1);
        border: 1px solid var(--teal-light);
        margin-bottom: 1.5rem;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(0, 128, 128, 0.15);
    }
    
    /* Button styling */
    .stButton button {
        background: linear-gradient(135deg, var(--primary-teal), var(--light-teal)) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(0, 128, 128, 0.2) !important;
        width: 100% !important;
    }
    
    .stButton button:hover {
        background: linear-gradient(135deg, var(--dark-teal), var(--primary-teal)) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(0, 128, 128, 0.3) !important;
    }
    
    /* Text area styling */
    .stTextArea textarea {
        border: 2px solid var(--teal-light) !important;
        border-radius: 8px !important;
        font-family: 'Courier New', monospace !important;
        transition: border-color 0.3s ease !important;
    }
    
    .stTextArea textarea:focus {
        border-color: var(--primary-teal) !important;
        box-shadow: 0 0 0 3px rgba(0, 128, 128, 0.1) !important;
    }
    
    /* Download button styling */
    .stDownloadButton button {
        background: linear-gradient(135deg, #28a745, #20c997) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(40, 167, 69, 0.2) !important;
        width: 100% !important;
    }
    
    .stDownloadButton button:hover {
        background: linear-gradient(135deg, #218838, #1ea976) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(40, 167, 69, 0.3) !important;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, var(--teal-bg), white) !important;
    }
    
    /* Info box styling */
    .info-box {
        background: linear-gradient(135deg, var(--teal-bg), var(--teal-light));
        border-left: 4px solid var(--primary-teal);
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    /* Success message styling */
    .stSuccess {
        background: linear-gradient(135deg, #d4edda, #c3e6cb) !important;
        border: 1px solid var(--light-teal) !important;
        border-radius: 8px !important;
    }
    
    /* Error message styling */
    .stError {
        border-radius: 8px !important;
    }
    
    /* Warning message styling */
    .stWarning {
        border-radius: 8px !important;
    }
    
    /* Metrics styling */
    .metric-card {
        background: var(--teal-bg);
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        border: 1px solid var(--teal-light);
    }
    
    /* Author info styling */
    .author-card {
        background: linear-gradient(135deg, var(--primary-teal), var(--light-teal));
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        margin: 1rem 0;
    }
    
    .author-avatar {
        background: rgba(255, 255, 255, 0.2);
        width: 80px;
        height: 80px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1rem auto;
        font-size: 2rem;
        font-weight: bold;
    }
    
    /* Feature list styling */
    .feature-list {
        list-style: none;
        padding: 0;
    }
    
    .feature-list li {
        padding: 0.5rem 0;
        position: relative;
        padding-left: 2rem;
    }
    
    .feature-list li:before {
        content: "✓";
        position: absolute;
        left: 0;
        color: var(--primary-teal);
        font-weight: bold;
        font-size: 1.2rem;
    }
    
    /* Responsive design */
    @media (max-width: 768px) {
        .main-header h1 {
            font-size: 2rem;
        }
        
        .card {
            padding: 1.5rem;
        }
    }
</style>
""", unsafe_allow_html=True)

def convert_markdown_to_word(markdown_text):
    """Chuyển đổi markdown thành file Word sử dụng pandoc"""
    try:
        # Tạo thư mục tạm thời
        temp_dir = tempfile.mkdtemp()
        
        # Đường dẫn files
        md_file = os.path.join(temp_dir, "input.md")
        docx_file = os.path.join(temp_dir, "output.docx")
        
        # Lưu markdown vào file
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(markdown_text)
        
        # Chạy pandoc để chuyển đổi
        pandoc_command = [
            "pandoc",
            md_file,
            "-o", docx_file,
            "--from", "markdown",
            "--to", "docx",
            "--standalone"
        ]
        
        result = subprocess.run(
            pandoc_command,
            check=True,
            capture_output=True,
            text=True
        )
        
        # Đọc file Word đã tạo
        with open(docx_file, "rb") as f:
            docx_data = f.read()
        
        # Dọn dẹp thư mục tạm
        shutil.rmtree(temp_dir)
        
        return {
            "success": True,
            "data": docx_data,
            "message": "Chuyển đổi thành công!"
        }
        
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "error": f"Lỗi pandoc: {e.stderr}",
            "message": "Chuyển đổi thất bại!"
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "message": "Có lỗi xảy ra!"
        }

def check_pandoc():
    """Kiểm tra xem pandoc có được cài đặt không"""
    try:
        result = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        return True, result.stdout.split('\n')[0]
    except:
        return False, "Pandoc chưa được cài đặt"

# Header chính
st.markdown("""
<div class="main-header">
    <h1>📝 Markdown to Word Converter</h1>
    <p>Chuyển đổi Markdown thành file Word một cách dễ dàng</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("""
    <div class="author-card">
        <div class="author-avatar">NP</div>
        <h3 style="margin: 0;">Nguyễn Hữu Phúc</h3>
        <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">Developer</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🎯 Tính năng")
    st.markdown("""
    <ul class="feature-list">
        <li>Dán markdown trực tiếp</li>
        <li>Chuyển đổi sang Word</li>
        <li>Hỗ trợ định dạng phong phú</li>
        <li>Giao diện thân thiện</li>
        <li>Xử lý nhanh chóng</li>
    </ul>
    """, unsafe_allow_html=True)
    
    # Kiểm tra pandoc
    pandoc_installed, pandoc_info = check_pandoc()
    
    if pandoc_installed:
        st.success(f"✅ {pandoc_info}")
    else:
        st.error(f"❌ {pandoc_info}")
    
    st.markdown("### 📞 Liên hệ")
    st.markdown("""
    - 📱 **Zalo:** [0985.692.879](https://zalo.me/0985692879)
    - 📘 **Facebook:** [nhphuclk](https://facebook.com/nhphuclk)  
    - 🌐 **Website:** [aiomtpremium.com](https://aiomtpremium.com)
    - 📺 **YouTube:** [@aiomtpremium](https://youtube.com/@aiomtpremium)
    """)

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("""
    <div class="card">
        <h3 style="color: var(--primary-teal); margin-top: 0;">✍️ Nhập nội dung Markdown</h3>
        <p style="color: #666;">Dán hoặc nhập nội dung Markdown của bạn vào ô bên dưới</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Text area để nhập markdown
    markdown_content = st.text_area(
        "Nội dung Markdown:",
        height=400,
        placeholder="""# Tiêu đề chính

## Tiêu đề phụ

Đây là đoạn văn bản **in đậm** và *in nghiêng*.

### Danh sách
- Mục 1
- Mục 2
- Mục 3

### Bảng
| Cột 1 | Cột 2 | Cột 3 |
|-------|-------|-------|
| A     | B     | C     |
| 1     | 2     | 3     |

> Đây là một trích dẫn

```python
print("Hello World!")
```
""",
        label_visibility="collapsed"
    )

with col2:
    st.markdown("""
    <div class="card">
        <h3 style="color: var(--primary-teal); margin-top: 0;">⚙️ Thao tác</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Thống kê nội dung
    if markdown_content:
        word_count = len(markdown_content.split())
        char_count = len(markdown_content)
        line_count = len(markdown_content.split('\n'))
        
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="margin: 0; color: var(--primary-teal);">📊 Thống kê</h4>
            <p style="margin: 0.5rem 0;"><strong>Từ:</strong> {word_count:,}</p>
            <p style="margin: 0.5rem 0;"><strong>Ký tự:</strong> {char_count:,}</p>
            <p style="margin: 0.5rem 0;"><strong>Dòng:</strong> {line_count:,}</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Nút chuyển đổi
    if st.button("🔄 Chuyển đổi sang Word", disabled=not markdown_content.strip()):
        if not pandoc_installed:
            st.error("❌ Pandoc chưa được cài đặt. Vui lòng cài đặt pandoc trước khi sử dụng.")
        else:
            with st.spinner("🔄 Đang chuyển đổi..."):
                result = convert_markdown_to_word(markdown_content)
                
                if result["success"]:
                    st.success("✅ " + result["message"])
                    
                    # Lưu kết quả vào session state
                    st.session_state.docx_data = result["data"]
                    st.session_state.conversion_success = True
                else:
                    st.error("❌ " + result["message"])
                    if "error" in result:
                        st.error(f"Chi tiết lỗi: {result['error']}")

    # Nút tải xuống (chỉ hiển thị sau khi chuyển đổi thành công)
    if hasattr(st.session_state, 'conversion_success') and st.session_state.conversion_success:
        if hasattr(st.session_state, 'docx_data'):
            st.download_button(
                label="⬇️ Tải file Word",
                data=st.session_state.docx_data,
                file_name="converted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Preview section
if markdown_content.strip():
    st.markdown("""
    <div class="card">
        <h3 style="color: var(--primary-teal); margin-top: 0;">👀 Xem trước</h3>
        <p style="color: #666;">Đây là cách nội dung sẽ hiển thị sau khi được định dạng</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Hiển thị preview markdown
    with st.expander("📋 Xem trước Markdown", expanded=True):
        st.markdown(markdown_content)

# Footer
st.markdown("""
<div style="text-align: center; padding: 2rem; margin-top: 3rem; background: linear-gradient(135deg, var(--teal-bg), white); border-radius: 12px;">
    <p style="color: var(--primary-teal); font-weight: 500; margin: 0;">
        💡 Hỗ trợ đầy đủ cú pháp Markdown • Chuyển đổi nhanh chóng • Giao diện thân thiện
    </p>
</div>
""", unsafe_allow_html=True)

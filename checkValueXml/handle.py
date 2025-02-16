

from rx.subject import Subject

# Tạo một Subject để phát sự kiện
button_click_subject = Subject()

def process_event(data):
    import main as m
    """Hàm này sẽ chạy khi sự kiện được kích hoạt."""
    print(f"📢 Nhận sự kiện: ")
    m.load_json_to_treeview(m.folderurl + "dataJson.json")
    

# Đăng ký hàm lắng nghe sự kiện
button_click_subject.subscribe(process_event)
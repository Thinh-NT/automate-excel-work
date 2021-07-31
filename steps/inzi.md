# CÁC BƯỚC TRONG BÁO CÁO INZI

- Bước 1:

  - Thêm 2 cột 'Material Type Type' và 'Procurement' ở bảng dữ liệu 2 sang 1.
  - Từ cột 'Movement Type' --> Tạo cột bên cạnh Account Code.
  - Từ cột 'Quantity' --> Tạo 2 cột bên cạnh 'Input', 'Output' (Quantity = Input - Output)

- Bước 2:

  - Từ 4 cột 'Account Code', 'Order Category', 'Material Type', 'Procurement' tách dữ liệu theo công thức:
    - Nhập vào:
      - 'Account Code' = 101 và 'Order Category' = 'Purchase order'
      - 'Account Code' = [602, 610, 623, 701, 720, 801, 809] và (
        'Material Type' = 'ROH' --> 'Procurement' = [F, X] |
        'Material Type' = 'HAWA' --> 'Procurement' = [F, X] |
        'Material Type' = 'HALB' --> 'Procurement' = F |
        'Material Type' = 'FERT' --> 'Procurement' = F |
        )
    - Xuất ra:
      - 'Account Code' = 102 và 'Order Category' = 'Purchase order'
      - 'Account Code' = [201, 261, 601, 609, 702, 712, 721, 803] và (
        'Material Type' = 'ROH' --> 'Procurement' = [F, X] |
        'Material Type' = 'HAWA' --> 'Procurement' = [F, X] |
        'Material Type' = 'HALB' --> 'Procurement' = F |
        'Material Type' = 'FERT' --> 'Procurement' = F |
        )
    - Thuyên chuyển:
      - 'Account Code' = [300 -> 401] và (
        'Material Type' = 'ROH' --> 'Procurement' = [F, X] |
        'Material Type' = 'HAWA' --> 'Procurement' = [F, X] |
        'Material Type' = 'HALB' --> 'Procurement' = F |
        'Material Type' = 'FERT' --> 'Procurement' = F |
        )
      - 'Account Code' = 555 và 'Material Type' = 'ROH' và 'Procurement' = [F, X]
        --> SCM RAW

- Bước 3:

  - Tương tự như bước 2:

    - Nhập vào:
      - 'Account Code' = 101 và 'Order Category' = 'Production Order'
      - 'Account Code' = [602, 610, 623, 701, 720, 801, 809] và (
        'Material Type' = 'HALB' --> 'Procurement' = [E, X]  
        )
    - Xuất ra:
      - 'Account Code' = 102 và 'Order Category' = 'Production Order'
      - 'Account Code' = [201, 261, 555, 601, 609, 702, 712, 721, 803] và (
        'Material Type' = 'HALB' --> 'Procurement' = [E, X]
        )
    - Thuyên chuyển:

      - 'Account Code' = [300 -> 401] và (
        'Material Type' = 'HALB' --> 'Procurement' = [E, X]  
        )

      - Với ('Account Code' = 261 và 'Material Type' = 'FERT') các kho bắt đầu bằng RW đổi thành Null (Không xài nữa)
        --> SCM WIP

- Bước 4:
  - Nhập vào:
    - 'Account Code' = 101 và 'Order Category' = 'Production order'
    - 'Account Code' = [602, 610, 623, 701, 720, 801, 809] và (
      'Material Type' = 'FERT' --> 'Procurement' = [E, X]  
       )
  - Xuất ra: - 'Account Code' = 102 và 'Order Category' = 'Production order' - 'Account Code' = [201, 261, 555, 601, 609, 702, 712, 721, 803] và (
    'Material Type' = 'FERT' --> 'Procurement' = [E, X]  
     )
  - Thuyên chuyển:
    - 'Account Code' = [300 -> 401] và (
      'Material Type' = 'FERT' --> 'Procurement' = [E, X]  
      )
      --> SCM FG
- Bước 5:
  - Tạo bảng mới với columns là các giá trị của 'Account Code', giá trị theo cột 'Quantity', giá trị của cùng 'Item Code' ('Material') thì được cộng dồn.
  - Cột 'IN' = Tổng các cột trong [101, 720, 801, 809]
  - Cột 'OUT' = Tổng các cột trong [102, 201, 261, 555, 601, 609, 721, 803] nếu có.

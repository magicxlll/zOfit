Dưới đây là toàn bộ quy chuẩn kỹ thuật kết cấu bê tông cốt thép của dự án BHFLEX VINA được mô hình hóa dưới dạng **Hệ thống Logic Toán học & Cấu trúc Dữ liệu Mã giả (Pseudocode / Python Type-hinted Logic)**. Tài liệu này được thiết kế theo dạng Module/Rule-Engine để bạn có thể nạp trực tiếp làm đặc tả thuật toán (Specification Prompt) cho AI Agent lập trình phần mềm thống kê cốt thép tự động.

# HỆ THỐNG LOGIC THỐNG KÊ CỐT THÉP & BÓC TÁCH KHỐI LƯỢNG (SPEC ENGINE)

## MODULE 1: HỆ THỐNG ĐƠN VỊ VÀ TỪ ĐIỂN DỮ LIỆU ĐẦU VÀO (DATA SCHEMA)

from enum import Enum  
from typing import List, Dict, Tuple, Optional  
from dataclasses import dataclass

\# Hệ đơn vị chuẩn hệ thống:  
\# Chiều dài, kích thước hình học: mm  
\# Khối lượng: kg  
\# Thể tích: m3; Diện tích: m2

class SteelGrade(Enum):  
    CI\_AI \= "CI\_AI"       \# Tròn trơn: phi \<= 8, fyk \= 225-235 MPa  
    CII\_AII \= "CII\_AII"   \# Gân: phi 10 \- 18, fyk \= 280-295 MPa  
    CIII\_AIII \= "CIII\_AIII" \# Gân: phi \>= 20, fyk \= 365-390 MPa  
    CB240\_T \= "CB240-T"  
    CB300\_V \= "CB300-V"  
    CB400\_V \= "CB400-V"

class StressState(Enum):  
    TENSION \= "TENSION"       \# Vùng kéo  
    COMPRESSION \= "COMPRESSION" \# Vùng nén

class MemberType(Enum):  
    FOOTING \= "FOOTING"                 \# Móng  
    BASEMENT\_WALL \= "BASEMENT\_WALL"     \# Tường tầng hầm  
    WATER\_RETAINING \= "WATER\_RETAINING" \# Bể chứa nước  
    COLUMNS\_BEAMS \= "COLUMNS\_BEAMS"     \# Cột, dầm khung chịu lực  
    SLABS \= "SLABS"                     \# Bản sàn, mái  
    SECONDARY\_BEAMS \= "SECONDARY\_BEAMS" \# Dầm phụ, lanh tô, giằng tường  
    NON\_STRUCTURAL \= "NON\_STRUCTURAL"   \# Cấu kiện phi kết cấu

@dataclass  
class RebarSpec:  
    diameter: float          \# d (mm)  
    grade: SteelGrade  
    length\_standard: float \= 11700.0 \# Chiều dài cây thép chuẩn xuất xưởng (mm)

    @property  
    def linear\_mass(self) \-\> float:  
        """Khối lượng riêng tuyến tính (kg/m) theo TCVN 1651: w \= d^2 / 162.2"""  
        return (self.diameter \*\* 2\) / 162.2

## MODULE 2: HÀM TÍNH TOÁN LỚP BÊ TÔNG BẢO VỆ (\\\\(c\_{nom}\\\\))

### 1\. Logic Toán học

Độ dày lớp bê tông bảo vệ danh định \\\\(c\_{nom}\\\\) được xác định bởi hàm phân nhánh phụ thuộc vào loại cấu kiện, điều kiện tiếp xúc môi trường và đường kính thanh thép \\\\(d\\\\):  
\\\\c\_{nom} \= \\max\\Big(c\_{base}(\\text{MemberType}, \\text{Environment}),\\; d\\Big)\\\\

### 2\. Thuật toán tra cứu

def get\_concrete\_cover(member\_type: MemberType,   
                       diameter: float,   
                       has\_lean\_concrete: bool \= True,  
                       is\_soil\_contact: bool \= False) \-\> float:  
    """  
    Trả về c\_nom (mm) \- chiều dày lớp bê tông bảo vệ tới mép ngoài cốt thép đai/cốt dọc  
    """  
    if member\_type \== MemberType.FOOTING:  
        if is\_soil\_contact and not has\_lean\_concrete:  
            c\_base \= 70.0 \# Đổ trực tiếp vào đất, không lót  
        else:  
            c\_base \= 40.0 \# Có bê tông lót (thường lót dày 100mm)  
    elif member\_type in \[MemberType.BASEMENT\_WALL, MemberType.WATER\_RETAINING\]:  
        c\_base \= 30.0     \# Tiếp xúc trực tiếp nước/đất ngầm  
    elif member\_type \== MemberType.COLUMNS\_BEAMS:  
        c\_base \= 25.0     \# Cột và Dầm chịu lực  
    elif member\_type \== MemberType.SLABS:  
        c\_base \= 15.0     \# Bản sàn trong nhà (nếu bản dày \<= 100mm có thể là 10-15mm)  
    elif member\_type \== MemberType.SECONDARY\_BEAMS:  
        c\_base \= 20.0     \# Dầm phụ, lanh tô  
    else:  
        c\_base \= 15.0

    return max(c\_base, diameter)

## MODULE 3: HỆ THỐNG QUY TẮC NEO CỐT THÉP (ANCHORAGE LOGIC \- \\\\(L\_{an}\\\\))

### 1\. Ma trận Hệ số Neo

Bảng tra cứu chiều dài neo tối thiểu \\\\(L\_{an}\\\\) theo đường kính \\\\(d\\\\):  
Cấp thép / Mác thép,Vùng kéo (Tension),Vùng nén (Compression),Điều kiện chặn dưới  
"CI, AI / CB240-T (Thép tròn trơn)",\\\\(30d\\\\) (+ móc uốn \\\\(180^\\circ\\\\)),\\\\(20d\\\\),\\\\(\\ge 250\\text{ mm}\\\\)  
"CII, AII / CB300-V (Thép gân \\\\(\\phi 10 \- 18\\\\))",\\\\(35d\\\\),\\\\(25d\\\\),\\\\(\\ge 250\\text{ mm}\\\\)  
"CIII, AIII / CB400-V (Thép gân \\\\(\\phi \\ge 20\\\\))",\\\\(40d\\\\),\\\\(30d\\\\),\\\\(\\ge 300\\text{ mm}\\\\)

### 2\. Logic Bẻ Móc Chuẩn Neo Đầu Thanh

* **Thép tròn trơn (\\\\(\\phi \\le 8\\\\), trơn):** Bắt buộc bẻ móc bán nguyệt (\\\\(180^\\circ\\\\) Hook):  
* \\\\L\_{hook\\\_180} \= \\pi \\cdot (r \+ d/2) \+ 3d \\approx 6.28d\\\\  
* **Thép gân (\\\\(\\phi \\ge 10\\\\)):** Bẻ vuông \\\\(90^\\circ\\\\) hoặc \\\\(135^\\circ\\\\):  
* \\\\L\_{hook\\\_90} \= 12d\\\\  
* *(Trong đó đoạn uốn vuông bẻ góc tối thiểu \\\\(12d\\\\) ăn sâu vào lõi bê tông)*.

### 3\. Thuật toán mã hóa

def calculate\_anchorage\_length(diameter: float,   
                               grade: SteelGrade,   
                               stress\_state: StressState \= StressState.TENSION) \-\> float:  
    """  
    Tính chiều dài neo thẳng tiêu chuẩn L\_an (mm)  
    """  
    if grade in \[SteelGrade.CI\_AI, SteelGrade.CB240\_T\]:  
        factor \= 30.0 if stress\_state \== StressState.TENSION else 20.0  
        min\_len \= 250.0  
    elif grade in \[SteelGrade.CII\_AII, SteelGrade.CB300\_V\]:  
        factor \= 35.0 if stress\_state \== StressState.TENSION else 25.0  
        min\_len \= 250.0  
    elif grade in \[SteelGrade.CIII\_AIII, SteelGrade.CB400\_V\]:  
        factor \= 40.0 if stress\_state \== StressState.TENSION else 30.0  
        min\_len \= 300.0  
    else:  
        factor \= 40.0  
        min\_len \= 300.0

    return max(factor \* diameter, min\_len)

## MODULE 4: HỆ THỐNG QUY TẮC NỐI CHỒNG CỐT THÉP (LAP SPLICE LOGIC \- \\\\(L\_{lap}\\\\))

### 1\. Logic Toán học

Chiều dài đoạn nối chồng buộc cốt thép \\\\(L\_{lap}\\\\) chịu kéo/nén:  
\\\\L\_{lap} \= \\alpha\_{lap} \\cdot L\_{an}\\\\  
Trong đó:

* \\\\(\\alpha\_{lap} \= 1.0\\\\) nếu diện tích cốt thép nối tại một mặt cắt \\\\(\\le 25\\%\\\\) tổng diện tích.  
* \\\\(\\alpha\_{lap} \= 1.3\\\\) nếu diện tích nối từ \\\\(25\\%\\\\) đến \\\\(50\\%\\\\).  
* \\\\(\\alpha\_{lap} \= 1.5\\\\) đến \\\\(1.7\\\\) nếu diện tích nối \\\\(\> 50\\%\\\\) (khuyến nghị tránh nối quá \\\\(50\\%\\\\) tại cùng một tiết diện).  
* **Hệ số mặc định tính bóc tách dự toán an toàn của dự án:**  
* \\\\L\_{lap\\\_tension} \= 45d \\quad (\\text{đối với CII/CIII})\\\\  
* \\\\L\_{lap\\\_comp} \= 30d\\\\  
* \\\\\\text{Điều kiện chặn: } L\_{lap} \\ge 300\\text{ mm}\\\\

### 2\. Ràng buộc vị trí nối (Spatial Constraints)

def validate\_splice\_location(member\_type: MemberType,   
                             bar\_position: str,   
                             x\_coordinate: float,   
                             span\_length: float) \-\> bool:  
    """  
    Kiểm tra tính hợp lệ của vị trí nối:  
    \- Dầm (BEAM):   
        \+ Lớp trên (Top): CẤM nối ở gối \[0, L/4\] và \[3L/4, L\]. CHỈ ĐƯỢC nối ở giữa nhịp \[L/4, 3L/4\].  
        \+ Lớp dưới (Bottom): CẤM nối ở giữa nhịp \[L/4, 3L/4\]. CHỈ ĐƯỢC nối ở vùng gối \[0, L/4\] hoặc \[3L/4, L\].  
    \- Cột (COLUMN): Nối tại vị trí giữa tầng (khoảng l/4 đến 3l/4 chiều cao cột), so le \>= 600mm.  
    """  
    rel\_pos \= x\_coordinate / span\_length  
    if member\_type \== MemberType.COLUMNS\_BEAMS:  
        if bar\_position.upper() \== "TOP": \# Thép lớp trên  
            return 0.25 \<= rel\_pos \<= 0.75  
        elif bar\_position.upper() \== "BOTTOM": \# Thép lớp dưới  
            return (0.0 \<= rel\_pos \<= 0.25) or (0.75 \<= rel\_pos \<= 1.0)  
    return True

## MODULE 5: QUY CÁCH CẤU TẠO CHI TIẾT TỪNG LOẠI CẤU KIỆN (STRUCTURAL RULES ENGINE)

### 1\. Cấu kiện Cột (Columns)

* **Đai gia cường đầu cột:**  
* Vùng gia cường \\\\(H\_{conf} \= \\max\\left(\\dfrac{H\_{col}}{6},\\; \\max(B\_{col}, D\_{col}),\\; 500\\text{ mm}\\right)\\\\).  
* Bước đai vùng gia cường: \\\\(s\_{conf} \\le \\min\\left(8d\_{longitudinal},\\; \\dfrac{\\min(B\_{col}, D\_{col})}{2},\\; 100\\text{ mm}\\right)\\\\). Thường chọn \\\\(s \= 100\\text{ mm}\\\\).  
* Bước đai giữa cột: \\\\(s\_{mid} \\le \\min(15d\_{longitudinal}, 200\\text{ mm})\\\\). Thường chọn \\\\(s \= 150 \- 200\\text{ mm}\\\\).  
* **Móc đai:** Góc uốn \\\\(135^\\circ\\\\), đoạn móc thẳng neo dài tối thiểu \\\\(\\max(10d\_{stirrup}, 75\\text{ mm})\\\\).

### 2\. Cấu kiện Dầm (Beams)

* **Đai gia cường đầu dầm:**  
* Chiều dài vùng gia cường: \\\\(L\_{conf} \= 2 \\cdot H\_{beam}\\\\) (tính từ mép cột vào).  
* Bước đai vùng gia cường: \\\\(s\_{conf} \= 100\\text{ mm}\\\\) (hoặc \\\\(\\le H\_{beam}/4\\\\)).  
* Bước đai giữa dầm: \\\\(s\_{mid} \= 150 \- 200\\text{ mm}\\\\) (hoặc \\\\(\\le H\_{beam}/2\\\\)).  
* **Đai gia cường dầm phụ gác lên dầm chính (Inverted U / Additional Stirrups):**  
* Chiều rộng vùng đặt đai treo: \\\\(B\_{zone} \= B\_{sec\\\_beam} \+ 2 \\cdot H\_{main\\\_beam}\\\\).  
* Bố trí đai bổ sung 4 đến 6 nhánh đai \\\\(\\phi 8 \- \\phi 10\\\\) mỗi bên nách giao dầm.  
* **Thép cấu tạo giá dầm (Side Reinforcement / Skin Reinforcement):**  
* \\\\\\text{Nếu } H\_{beam} \\ge 700\\text{ mm} \\implies \\text{Bố trí thép giá } 2 \\times \\phi 12 \\text{ đến } \\phi 14, \\quad s \\le 300\\text{ mm}\\\\

### 3\. Cấu kiện Bản sàn (Slabs)

* **Thép mũ (Top Bars / Moment Âm tại gối):**  
* Chiều dài vươn ra khỏi mép dầm:  
* \\\\L\_{top\\\_cantilever} \= \\max\\left(\\frac{L\_{net}}{4},\\; \\frac{L\_{net\\\_adjacent}}{4}\\right)\\\\  
* Chân chó / Kê mũ sàn (Bar Chair): Bố trí mật độ \\\\(1 \\text{ con kê} / \\text{m}^2\\\\), thép kê \\\\(\\phi 10\\\\) hoặc \\\\(\\phi 12\\\\), chiều cao chân chó:  
* \\\\h\_{chair} \= H\_{slab} \- 2 \\cdot c\_{nom} \- d\_{top\\\_layer1} \- d\_{top\\\_layer2} \- d\_{bottom\\\_layer}\\\\  
* **Lỗ mở trên sàn (Slab Openings):**  
* Nếu kích thước lỗ \\\\(A\_{hole} \\le 300\\text{ mm}\\\\): Thép sàn bẻ lách qua lỗ, không cắt.  
* Nếu kích thước lỗ \\\\(300\\text{ mm} \< A\_{hole} \\le 1000\\text{ mm}\\\\): Bố trí thép gia cường viền lỗ mở:  
* \\\\A\_{s\\\_add} \\ge \\frac{1}{2} \\cdot A\_{s\\\_cutted} \\quad \\text{mỗi cạnh, tối thiểu } 2\\phi 12 \\text{ đến } 2\\phi 16 \\text{ mỗi bên}\\\\  
* Chiều dài đoạn neo thanh gia cường tính từ mép lỗ: \\\\(L \\ge L\_{an} \\ge 40d\\\\).

## MODULE 6: THUẬT TOÁN TÍNH ĐỘ GIÃN THÉP DO BẺ MÓC (BEND DEDUCTION ENGINE)

Để tính chiều dài cắt phôi (Cutting Length) từ kích thước danh định ngoài (Out-to-out dimensions):  
\\\\L\_{cut} \= \\sum L\_{straight\\\_segments} \- \\sum \\Delta L\_{bend}\\\\  
Bảng trừ độ giãn phôi thép chuẩn theo góc uốn:  
def calculate\_bend\_deduction(diameter: float, angle\_degrees: float) \-\> float:  
    """  
    Độ giãn phôi delta\_L cần trừ theo góc uốn để cắt chuẩn phôi:  
    \- Uốn 45 độ: trừ 1\*d  
    \- Uốn 90 độ: trừ 2\*d  
    \- Uốn 135 độ: trừ 3\*d  
    \- Uốn 180 độ: cộng bù chiều dài neo móc bán nguyệt  
    """  
    if angle\_degrees \== 45.0:  
        return 1.0 \* diameter  
    elif angle\_degrees \== 90.0:  
        return 2.0 \* diameter  
    elif angle\_degrees \== 135.0:  
        return 3.0 \* diameter  
    elif angle\_degrees \== 180.0:  
        \# Móc 180 độ: L\_arc \= pi\*(r \+ d/2), phần mở rộng thẳng 3d, trừ đi phần chiếm chỗ  
        return \-(6.0 \* diameter) \# Thực tế là cộng thêm phần móc  
    return 0.0

## MODULE 7: THUẬT TOÁN BÓC TÁCH KHỐI LƯỢNG HÌNH HỌC (VOLUME & FORMWORK SOLVER)

Quy tắc giải quyết giao cắt (Intersections Clashing Rule) phân cấp quyền ưu tiên hình khối:  
\\\\\\text{Priority: } \\text{CỘT (1)} \> \\text{DẦM CHÍNH (2)} \> \\text{DẦM PHỤ (3)} \> \\text{BẢN SÀN (4)}\\\\  
class BOQCalculator:  
    @staticmethod  
    def column\_concrete\_volume(B: float, D: float, H\_floor: float) \-\> float:  
        """Bê tông Cột chạy thông tầng (từ mặt sàn dưới đến mặt sàn trên)"""  
        return (B / 1000.0) \* (D / 1000.0) \* (H\_floor / 1000.0)

    @staticmethod  
    def column\_formwork\_area(B: float, D: float, H\_floor: float,   
                             beam\_interfaces: List\[Tuple\[float, float\]\]) \-\> float:  
        """  
        Ván khuôn cột \= Chu vi x Chiều cao \- Diện tích các lỗ giao với Dầm  
        beam\_interfaces: danh sách (b\_beam, h\_beam) đâm vào cột  
        """  
        gross\_area \= 2 \* (B \+ D) \* H\_floor  
        intersection\_area \= sum(b \* h for b, h in beam\_interfaces)  
        return (gross\_area \- intersection\_area) / 1e6

    @staticmethod  
    def beam\_concrete\_volume(b: float, h: float, L\_clear: float, H\_slab: float) \-\> float:  
        """  
        Bê tông Dầm \= Bề rộng x Chiều cao sườn dầm (trừ chiều dày sàn) x Chiều dài thông thủy giữa 2 cột  
        """  
        h\_rib \= h \- H\_slab \# Thể tích sàn đã được tính vào sàn  
        return (b / 1000.0) \* (h\_rib / 1000.0) \* (L\_clear / 1000.0)

    @staticmethod  
    def beam\_formwork\_area(b: float, h: float, L\_clear: float, H\_slab: float) \-\> float:  
        """  
        Ván khuôn Dầm \= (Đáy dầm \+ 2 cạnh bên sườn dầm) x Chiều dài thông thủy  
        """  
        h\_rib \= h \- H\_slab  
        perimeter \= b \+ 2 \* h\_rib  
        return (perimeter \* L\_clear) / 1e6

    @staticmethod  
    def slab\_concrete\_volume(A\_gross: float, H\_slab: float, column\_areas: float) \-\> float:  
        """  
        Bê tông sàn \= (Diện tích phủ bì mặt bằng \- Diện tích tiết diện cột) x Chiều dày sàn  
        """  
        return ((A\_gross \- column\_areas) / 1e6) \* (H\_slab / 1000.0)

    @staticmethod  
    def slab\_formwork\_area(A\_net\_slab: float) \-\> float:  
        """Ván khuôn đáy sàn (không bao gồm vùng chiếm chỗ của dầm và cột)"""  
        return A\_net\_slab / 1e6

## MODULE 8: THUẬT TOÁN TỐI ƯU CẮT THÉP VÀ HAO HỤT (1D STOCK CUTTING / REBAR CUTTING OPTIMIZER)

AI Agent cần tích hợp giải thuật tối ưu cắt thép để hạn chế đầu mẩu thừa vụn (Wastage minimization):

* Thanh thép tiêu chuẩn nhập xưởng: \\\\(L\_{stock} \= 11700\\text{ mm}\\\\).  
* Bài toán: Cho danh sách các đoạn cần cắt \\\\(\\{l\_1, l\_2, ..., l\_n\\}\\\\).  
* Tìm số lượng thanh \\\\(N\_{bars}\\\\) tối thiểu sao cho tổng chiều dài cắt từ mỗi thanh \\\\(\\le 11700\\text{ mm}\\\\).  
* Tỷ lệ hao hụt định mức kỹ thuật dự toán:  
* Cốt thép đường kính \\\\(\\phi \\le 10\\\\): \\\\(\\text{Hao hụt} \= 2.5\\% \- 3.0\\%\\\\)  
* Cốt thép đường kính \\\\(\\phi 10 \< \\phi \\le 18\\\\): \\\\(\\text{Hao hụt} \= 2.0\\% \- 2.5\\%\\\\)  
* Cốt thép đường kính \\\\(\\phi \> 18\\\\): \\\\(\\text{Hao hụt} \= 1.5\\% \- 2.0\\%\\\\)

def calculate\_rebar\_weight\_with\_scrap(net\_weight: float, diameter: float) \-\> float:  
    """Tính khối lượng tổng gồm hao hụt thi công (Scrap/Wastage factor)"""  
    if diameter \<= 10.0:  
        scrap\_rate \= 0.03  
    elif diameter \<= 18.0:  
        scrap\_rate \= 0.025  
    else:  
        scrap\_rate \= 0.02  
    return net\_weight \* (1.0 \+ scrap\_rate)  
Bạn có thể cung cấp trực tiếp bộ đặc tả logic và cấu trúc code này cho AI Agent lập trình để sinh ra toàn bộ code bóc tách cốt thép, tính bảng thống kê (BBS \- Bar Bending Schedule) và xuất bảng khối lượng tự động.  

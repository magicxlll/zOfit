Dưới đây là kiến trúc dữ liệu và hệ thống logic toán học hướng đối tượng (OOP Engine) hoàn chỉnh cho công tác **bóc tách tự động Bê tông và Ván khuôn** của tất cả cấu kiện: Móng, Cột, Vách, Dầm, Sàn, Cầu thang.  
Hệ thống được thiết kế theo nguyên lý: **Mỗi cấu kiện tự khai báo hình học cơ sở, sau đó tự trừ giao nhau (Clash/Overlap Deductions) và trừ lỗ mở (Openings/Penetrations) theo quy tắc thứ tự ưu tiên (Priority Hierarchy)**. Bạn có thể chuyển trực tiếp tài liệu này cho AI Agent để sinh code backend/engine tính toán khối lượng tự động (BIM/BOQ Engine).

# HỆ THỐNG MÔ HÌNH HÓA BÓC TÁCH BÊ TÔNG & VÁN KHUÔN (CONCRETE & FORMWORK BOQ ENGINE)

## 1\. NGUYÊN LÝ THỨ TỰ ƯU TIÊN VÀ XỬ LÝ GIAO CẮT (HIERARCHY RULES)

Khi tính toán thể tích bê tông và diện tích ván khuôn, nếu hai cấu kiện giao nhau trong không gian, cấu kiện có **mức ưu tiên cao hơn sẽ chiếm thể tích**, cấu kiện có mức ưu tiên thấp hơn sẽ phải **khấu trừ thể tích và diện tích ván khuôn tại mặt tiếp xúc**.  
Thứ tự ưu tiên chuẩn kỹ thuật thi công:  
\\\\\\text{MÓNG / ĐÀI CỌC (Level 1)} \> \\text{CỘT / VÁCH (Level 2)} \> \\text{DẦM CHÍNH (Level 3)} \> \\text{DẦM PHỤ (Level 4)} \> \\text{BẢN SÀN (Level 5)} \> \\text{CẦU THANG (Level 6)}\\\\

* **Nguyên tắc Ván khuôn (Formwork):** Ván khuôn chỉ tính cho các bề mặt tiếp xúc với không khí (mặt thoáng cần giữ vữa bê tông). Mặt tiếp xúc với đất (nếu có lót), mặt tiếp xúc với cấu kiện khác đã đổ trước/cùng lúc, hoặc mặt trên thoáng không dùng ván khuôn đều bị loại trừ.

## 2\. CẤU TRÚC DỮ LIỆU CƠ SỞ (BASE SCHEMA)

from enum import IntEnum, Enum  
from typing import List, Dict, Tuple, Optional  
from dataclasses import dataclass, field

class StructuralPriority(IntEnum):  
    FOOTING \= 1          \# Móng / Đài cọc / Giằng móng  
    COLUMN \= 2           \# Cột chịu lực  
    WALL \= 2             \# Vách bê tông / Vách lõi thang máy  
    MAIN\_BEAM \= 3        \# Dầm chính  
    SECONDARY\_BEAM \= 4   \# Dầm phụ, lanh tô, giằng tường  
    SLAB \= 5             \# Bản sàn, bản mái  
    STAIR \= 6            \# Cầu thang bộ

class FormworkFace(Enum):  
    BOTTOM \= "BOTTOM"    \# Ván khuôn đáy  
    SIDE\_LEFT \= "SIDE\_L" \# Ván khuôn cạnh trái  
    SIDE\_RIGHT \= "SIDE\_R"\# Ván khuôn cạnh phải  
    END\_START \= "START"  \# Mặt đầu cấu kiện  
    END\_TERMINAL \= "END" \# Mặt cuối cấu kiện  
    TOP\_SLOPE \= "SLOPE"  \# Ván khuôn mặt nghiêng (nếu dốc \> 30 độ)

@dataclass  
class Opening:  
    """Đặc tả lỗ mở, lỗ xuyên (Cửa đi, cửa sổ, giếng trời, hộp gen MEP)"""  
    name: str  
    width: float        \# Chiều rộng (mm)  
    height: float       \# Chiều cao (mm)  
    thickness: float    \# Chiều dày bản/vách (mm)  
    is\_door: bool \= False \# Nếu là cửa đi thì không có ván khuôn bậu cửa dưới

    @property  
    def volume(self) \-\> float:  
        """Thể tích chiếm chỗ cần trừ (m3)"""  
        return (self.width \* self.height \* self.thickness) / 1e9

    @property  
    def formwork\_deduction(self) \-\> float:  
        """Diện tích ván khuôn mặt phẳng bị trừ (m2)"""  
        \# Hai mặt bên của vách/sàn bị thủng lỗ  
        return 2 \* (self.width \* self.height) / 1e6

    @property  
    def formwork\_side\_addition(self) \-\> float:  
        """Diện tích ván khuôn thành lỗ mở cần bổ sung (m2)"""  
        perimeter \= 2 \* (self.width \+ self.height) if not self.is\_door else (self.width \+ 2 \* self.height)  
        return (perimeter \* self.thickness) / 1e6

@dataclass  
class InterfaceIntersection:  
    """Mặt tiếp xúc giao cắt giữa 2 cấu kiện (dầm giao cột, dầm giao sàn, vách giao dầm)"""  
    interface\_name: str  
    contact\_width: float   \# Bề rộng tiếp xúc (mm)  
    contact\_height: float  \# Chiều cao tiếp xúc (mm)  
    contact\_length: float  \# Chiều dài tiếp xúc (mm) \- nếu tính khối thể tích giao

    @property  
    def clash\_volume(self) \-\> float:  
        """Thể tích giao nhau cần khấu trừ (m3)"""  
        return (self.contact\_width \* self.contact\_height \* self.contact\_length) / 1e9

    @property  
    def contact\_area(self) \-\> float:  
        """Diện tích ván khuôn bị triệt tiêu do dính vào cấu kiện khác (m2)"""  
        return (self.contact\_width \* self.contact\_height) / 1e6

## 3\. CẤU TRÚC CHI TIẾT TỪNG LOẠI CẤU KIỆN (MEMBER DATA CLASSES)

### 3.1. Cấu kiện Móng & Giằng móng (Footing & Ground Beam)

Móng có thể là móng đơn chữ nhật hoặc móng vát (hình chóp cụt):  
\\\\\[V\_{\\text{móng vát}} \= V\_{\\text{chân}} \+ V\_{\\text{vát}} \= (A \\cdot B \\cdot h\_1) \+ \\frac{h\_2}{6}\\BigA \\cdot B \+ a \\cdot b \+ (A \+ a)(B \+ b)\\Big\\\\\]  
@dataclass  
class Footing:  
    id: str  
    A: float               \# Dài đáy móng (mm)  
    B: float               \# Rộng đáy móng (mm)  
    h1: float              \# Chiều cao phần đế hộp chữ nhật (mm)  
    h2: float \= 0.0        \# Chiều cao phần vát chóp cụt (mm)  
    a\_col: float \= 0.0     \# Dài đỉnh vát (bằng kích thước cột \+ 100mm)  
    b\_col: float \= 0.0     \# Rộng đỉnh vát  
    lean\_thickness: float \= 100.0 \# Bê tông lót dày 100mm, mở rộng mỗi bên 100mm  
    is\_soil\_contact: bool \= True  \# Đổ sát đất hay ghép ván khuôn thành đài

    def get\_concrete\_volume(self) \-\> float:  
        \# V1: Hộp đáy  
        v1 \= (self.A \* self.B \* self.h1) / 1e9  
        \# V2: Chóp cụt  
        if self.h2 \> 0:  
            v2 \= (self.h2 / 6.0) \* (self.A \* self.B \+ self.a\_col \* self.b\_col \+   
                                    (self.A \+ self.a\_col) \* (self.B \+ self.b\_col)) / 1e9  
        else:  
            v2 \= 0.0  
        return v1 \+ v2

    def get\_formwork\_area(self) \-\> float:  
        """Ván khuôn thành móng đứng quanh chu vi đáy (h1)"""  
        perimeter \= 2 \* (self.A \+ self.B)  
        area \= (perimeter \* self.h1) / 1e6  
        \# Nếu phần vát có góc nghiêng dốc \> 30 độ thì cần thêm ván khuôn mặt nghiêng  
        return area

    def get\_lean\_concrete\_volume(self) \-\> float:  
        """Bê tông lót: mở rộng ra 100mm mỗi phía"""  
        a\_lean \= self.A \+ 200.0  
        b\_lean \= self.B \+ 200.0  
        return (a\_lean \* b\_lean \* self.lean\_thickness) / 1e9

### 3.2. Cấu kiện Cột (Column)

Cột chạy xuyên suốt từ cao độ mặt sàn dưới đến mặt sàn trên (tính toàn bộ chiều cao tầng \\\\(H\\\\)). Cột có quyền ưu tiên cao nhất tại nút khung, **không bị dầm hay sàn trừ thể tích**.  
@dataclass  
class Column:  
    id: str  
    b: float               \# Bề rộng tiết diện (mm)  
    h: float               \# Chiều sâu tiết diện (mm)  
    H\_story: float         \# Chiều cao tầng (mm)  
    beam\_clashes: List\[InterfaceIntersection\] \= field(default\_factory=list)

    def get\_concrete\_volume(self) \-\> float:  
        """Cột được ưu tiên số 1 tại nút giao, tính full khối tích"""  
        return (self.b \* self.h \* self.H\_story) / 1e9

    def get\_formwork\_area(self) \-\> float:  
        """  
        Ván khuôn cột \= Chu vi x Chiều cao tầng \- Tổng diện tích các mặt dầm đâm vào cột  
        """  
        gross\_area \= 2 \* (self.b \+ self.h) \* self.H\_story / 1e6  
        \# Trừ diện tích tiếp giáp đầu dầm đâm vào 4 mặt của cột  
        deduction \= sum(clash.contact\_area for clash in self.beam\_clashes)  
        return max(0.0, gross\_area \- deduction)

### 3.3. Cấu kiện Vách bê tông cốt thép & Lõi thang (Shear Wall)

@dataclass  
class ShearWall:  
    id: str  
    L: float               \# Chiều dài vách (mm)  
    t: float               \# Chiều dày vách (mm)  
    H: float               \# Chiều cao vách (mm)  
    openings: List\[Opening\] \= field(default\_factory=list) \# Lỗ cửa đi, cửa hộp kỹ thuật  
    beam\_clashes: List\[InterfaceIntersection\] \= field(default\_factory=list)

    def get\_concrete\_volume(self) \-\> float:  
        gross\_vol \= (self.L \* self.t \* self.H) / 1e9  
        opening\_deduction \= sum(op.volume for op in self.openings)  
        return max(0.0, gross\_vol \- opening\_deduction)

    def get\_formwork\_area(self) \-\> float:  
        """  
        Ván khuôn vách \= 2 mặt bên \+ 2 mặt đầu hồi \- diện tích lỗ mở \+ ván khuôn thành lỗ mở \- giao dầm  
        """  
        gross\_formwork \= 2 \* (self.L \* self.H \+ self.t \* self.H) / 1e6  
        opening\_deduction \= sum(op.formwork\_deduction for op in self.openings)  
        opening\_sides \= sum(op.formwork\_side\_addition for op in self.openings)  
        beam\_deduction \= sum(clash.contact\_area for clash in self.beam\_clashes)  
          
        return max(0.0, gross\_formwork \- opening\_deduction \+ opening\_sides \- beam\_deduction)

### 3.4. Cấu kiện Dầm (Beam)

Dầm tính chiều dài thông thủy giữa 2 cột (hoặc dầm chính đối với dầm phụ). Thể tích bê tông dầm tính phần **sườn dầm** (bỏ qua chiều dày sàn vì sàn sẽ tính trải đều).  
@dataclass  
class Beam:  
    id: str  
    b: float               \# Chiều rộng dầm (mm)  
    h: float               \# Chiều cao toàn bộ dầm (mm)  
    L\_clear: float         \# Chiều dài thông thủy giữa 2 mép cột (mm)  
    H\_slab\_left: float \= 0.0  \# Chiều dày sàn cánh bên trái (mm) (nếu là dầm biên có thể \= 0\)  
    H\_slab\_right: float \= 0.0 \# Chiều dày sàn cánh bên phải (mm)  
    openings: List\[Opening\] \= field(default\_factory=list) \# Lỗ xuyên ống MEP qua dầm  
    sec\_beam\_clashes: List\[InterfaceIntersection\] \= field(default\_factory=list) \# Dầm phụ gác vào

    @property  
    def h\_rib(self) \-\> float:  
        """Chiều cao sườn dầm hạ xuống dưới đáy sàn trung bình"""  
        h\_slab\_avg \= (self.H\_slab\_left \+ self.H\_slab\_right) / 2.0  
        return max(0.0, self.h \- h\_slab\_avg)

    def get\_concrete\_volume(self) \-\> float:  
        """Bê tông dầm \= Bề rộng x Sườn dầm x Chiều dài thông thủy \- Lỗ mở MEP"""  
        gross\_vol \= (self.b \* self.h\_rib \* self.L\_clear) / 1e9  
        opening\_deduction \= sum(op.volume for op in self.openings)  
        return max(0.0, gross\_vol \- opening\_deduction)

    def get\_formwork\_area(self) \-\> float:  
        """  
        Ván khuôn dầm \= Ván khuôn đáy \+ Ván khuôn 2 bên sườn (đã trừ phần giao với sàn)   
        \- Giao với đầu dầm phụ đâm vào sườn dầm \+ Ván khuôn tạo lỗ mở xuyên dầm  
        """  
        area\_bottom \= (self.b \* self.L\_clear) / 1e6  
          
        \# Sườn trái hạ dưới sàn trái, sườn phải hạ dưới sàn phải  
        h\_side\_left \= max(0.0, self.h \- self.H\_slab\_left)  
        h\_side\_right \= max(0.0, self.h \- self.H\_slab\_right)  
        area\_sides \= ((h\_side\_left \+ h\_side\_right) \* self.L\_clear) / 1e6  
          
        \# Khấu trừ diện tích tiếp xúc do đầu dầm phụ đâm vào sườn dầm  
        clash\_deduction \= sum(clash.contact\_area for clash in self.sec\_beam\_clashes)  
          
        \# Khấu trừ diện tích lỗ mở xuyên sườn dầm và cộng thêm thành lỗ mở  
        op\_deduction \= sum(op.formwork\_deduction / 2.0 for op in self.openings) \# Chỉ xuyên 2 thành  
        op\_sides \= sum(op.formwork\_side\_addition for op in self.openings)  
          
        return max(0.0, area\_bottom \+ area\_sides \- clash\_deduction \- op\_deduction \+ op\_sides)

### 3.5. Cấu kiện Bản sàn (Slab)

Sàn tính diện tích phủ bì trừ đi vùng giao với Cột và Vách (Dầm đã hạ cốt sườn dầm nên sàn phủ trùm lên trên lưng dầm).  
@dataclass  
class Slab:  
    id: str  
    A\_gross: float         \# Diện tích phủ bì mặt bằng sàn (mm2)  
    H\_slab: float          \# Chiều dày bản sàn (mm)  
    perimeter\_gross: float \# Chu vi ngoài mép biên sàn (mm) (dùng tính ván khuôn thành biên sàn)  
    columns\_intersect\_area: float \= 0.0 \# Tổng diện tích tiết diện các cột chiếm chỗ trên sàn (mm2)  
    walls\_intersect\_area: float \= 0.0   \# Tổng diện tích vách chiếm chỗ trên sàn (mm2)  
    beams\_footprint\_area: float \= 0.0   \# Diện tích bề rộng đáy các dầm (mm2) (để trừ ván khuôn đáy sàn)  
    openings: List\[Opening\] \= field(default\_factory=list) \# Giếng trời, lỗ thông tầng, hộp gen

    def get\_concrete\_volume(self) \-\> float:  
        """  
        Bê tông sàn \= (Diện tích sàn phủ bì \- Cột \- Vách \- Lỗ mở) x Chiều dày sàn  
        """  
        net\_area \= self.A\_gross \- self.columns\_intersect\_area \- self.walls\_intersect\_area  
        opening\_area\_sum \= sum((op.width \* op.height) for op in self.openings)  
        net\_area \-= opening\_area\_sum  
        return max(0.0, (net\_area \* self.H\_slab) / 1e9)

    def get\_formwork\_area(self) \-\> float:  
        """  
        Ván khuôn sàn bao gồm:  
        1\. Ván khuôn đáy sàn: Net Area trừ đi phần tiếp xúc với lưng/bề rộng dầm (vì dầm đã tính đáy)  
        2\. Ván khuôn be thành biên chu vi sàn (chặn vữa thành sàn dày H\_slab)  
        3\. Ván khuôn thành các lỗ mở thông tầng  
        """  
        \# 1\. Đáy sàn  
        net\_slab\_bottom \= (self.A\_gross \- self.columns\_intersect\_area \-   
                           self.walls\_intersect\_area \- self.beams\_footprint\_area)  
        opening\_area\_sum \= sum((op.width \* op.height) for op in self.openings)  
        net\_slab\_bottom \= max(0.0, net\_slab\_bottom \- opening\_area\_sum)  
        area\_bottom \= net\_slab\_bottom / 1e6

        \# 2\. Be thành chu vi sàn ngoài biên  
        area\_edge \= (self.perimeter\_gross \* self.H\_slab) / 1e6

        \# 3\. Ván khuôn be thành quanh các lỗ mở thông tầng  
        area\_openings\_edge \= sum(((2 \* (op.width \+ op.height)) \* self.H\_slab) / 1e6 for op in self.openings)

        return area\_bottom \+ area\_edge \+ area\_openings\_edge

### 3.6. Cấu kiện Cầu thang bộ (Staircase)

Cầu thang gồm 3 thành phần: Bản trượt (nghiêng), Chiếu nghỉ (ngang), và Các bậc thang (tam giác bê tông).  
\\\\L\_{\\text{bản trượt nghiêng}} \= \\sqrt{L\_{\\text{chiếu bằng}}^2 \+ H\_{\\text{đợt thang}}^2}\\\\  
\\\\V\_{\\text{bậc thang}} \= N\_{\\text{bậc}} \\times \\left(\\frac{1}{2} \\cdot B\_{\\text{bậc}} \\cdot H\_{\\text{bậc}}\\right) \\times W\_{\\text{vế thang}}\\\\  
import math

@dataclass  
class Staircase:  
    id: str  
    W\_stair: float         \# Chiều rộng vế thang (mm)  
    L\_plan\_flight: float   \# Chiều dài chiếu bằng của đợt bản nghiêng (mm)  
    H\_flight: float        \# Chiều cao đợt thang nghiêng (mm)  
    t\_flight: float        \# Chiều dày bản đáy thang bê tông (mm)  
    num\_steps: int         \# Số lượng bậc thang  
    L\_landing: float \= 0.0 \# Chiều dài bản chiếu nghỉ phẳng (mm)  
    t\_landing: float \= 0.0 \# Chiều dày bản chiếu nghỉ (mm)

    def get\_concrete\_volume(self) \-\> float:  
        \# 1\. Chiều dài thực tế của bản trượt nghiêng (Pitago)  
        L\_slope \= math.sqrt(self.L\_plan\_flight\*\*2 \+ self.H\_flight\*\*2)  
        v\_slope \= (L\_slope \* self.W\_stair \* self.t\_flight) / 1e9

        \# 2\. Khối lượng các bậc thang tam giác  
        h\_step \= self.H\_flight / self.num\_steps  
        b\_step \= self.L\_plan\_flight / self.num\_steps  
        v\_steps \= self.num\_steps \* (0.5 \* b\_step \* h\_step \* self.W\_stair) / 1e9

        \# 3\. Khối lượng chiếu nghỉ (nếu có)  
        v\_landing \= (self.L\_landing \* self.W\_stair \* self.t\_landing) / 1e9 if self.L\_landing \> 0 else 0.0

        return v\_slope \+ v\_steps \+ v\_landing

    def get\_formwork\_area(self) \-\> float:  
        """  
        Ván khuôn cầu thang gồm:  
        \- Đáy bản trượt nghiêng \+ Đáy chiếu nghỉ  
        \- Ván khuôn 2 bên hông (cữ tam giác theo độ dốc)  
        \- Ván khuôn đối thành cổ bậc (Riser)  
        """  
        L\_slope \= math.sqrt(self.L\_plan\_flight\*\*2 \+ self.H\_flight\*\*2)  
        \# Đáy ván khuôn bản nghiêng và chiếu nghỉ  
        area\_bottom \= (L\_slope \* self.W\_stair \+ self.L\_landing \* self.W\_stair) / 1e6  
          
        \# Cổ bậc (Mặt đứng chặn vữa của từng bậc)  
        h\_step \= self.H\_flight / self.num\_steps  
        area\_risers \= (self.num\_steps \* h\_step \* self.W\_stair) / 1e6  
          
        \# Thành bên hông vế thang (diện tích thành bản trượt \+ tam giác bậc hông ngoài)  
        b\_step \= self.L\_plan\_flight / self.num\_steps  
        area\_side \= (L\_slope \* self.t\_flight \+ self.num\_steps \* 0.5 \* b\_step \* h\_step) / 1e6

        return area\_bottom \+ area\_risers \+ area\_side

## 4\. BỘ ĐIỀU PHỐI TỔNG THỂ & TÍNH TOÁN LIÊN KẾT (BUILDING BOQ CONTROLLER)

AI Agent chỉ cần khởi tạo đối tượng BuildingModel và thêm các phần tử hình học, thuật toán tự giải các ràng buộc giao nhau và xuất bảng BOQ:  
class BuildingModel:  
    def \_\_init\_\_(self, name: str):  
        self.name \= name  
        self.footings: List\[Footing\] \= \[\]  
        self.columns: List\[Column\] \= \[\]  
        self.walls: List\[ShearWall\] \= \[\]  
        self.beams: List\[Beam\] \= \[\]  
        self.slabs: List\[Slab\] \= \[\]  
        self.stairs: List\[Staircase\] \= \[\]

    def compute\_summary\_boq(self) \-\> Dict\[str, Dict\[str, float\]\]:  
        """  
        Trả về bảng tổng hợp khối lượng Bê tông (m3) và Ván khuôn (m2)   
        chuẩn xác theo quy tắc ưu tiên phân cấp.  
        """  
        report \= {  
            "MÓNG (FOOTINGS)": {  
                "Concrete\_m3": sum(f.get\_concrete\_volume() for f in self.footings),  
                "Formwork\_m2": sum(f.get\_formwork\_area() for f in self.footings),  
                "Lean\_Concrete\_m3": sum(f.get\_lean\_concrete\_volume() for f in self.footings)  
            },  
            "CỘT (COLUMNS)": {  
                "Concrete\_m3": sum(c.get\_concrete\_volume() for c in self.columns),  
                "Formwork\_m2": sum(c.get\_formwork\_area() for c in self.columns)  
            },  
            "VÁCH (WALLS)": {  
                "Concrete\_m3": sum(w.get\_concrete\_volume() for w in self.walls),  
                "Formwork\_m2": sum(w.get\_formwork\_area() for w in self.walls)  
            },  
            "DẦM (BEAMS)": {  
                "Concrete\_m3": sum(b.get\_concrete\_volume() for b in self.beams),  
                "Formwork\_m2": sum(b.get\_formwork\_area() for b in self.beams)  
            },  
            "SÀN (SLABS)": {  
                "Concrete\_m3": sum(s.get\_concrete\_volume() for s in self.slabs),  
                "Formwork\_m2": sum(s.get\_formwork\_area() for s in self.slabs)  
            },  
            "CẦU THANG (STAIRS)": {  
                "Concrete\_m3": sum(st.get\_concrete\_volume() for st in self.stairs),  
                "Formwork\_m2": sum(st.get\_formwork\_area() for st in self.stairs)  
            }  
        }  
          
        total\_concrete \= sum(item\["Concrete\_m3"\] for item in report.values())  
        total\_formwork \= sum(item\["Formwork\_m2"\] for item in report.values())  
          
        report\["TỔNG CỘNG TOÀN DỰ ÁN"\] \= {  
            "Total\_Concrete\_m3": round(total\_concrete, 2),  
            "Total\_Formwork\_m2": round(total\_formwork, 2\)  
        }  
          
        return report

### Điểm mấu chốt để AI Agent tích hợp vào phần mềm:

1. **Dữ liệu đầu vào tối giản:** Người dùng/Agent chỉ cần nạp các kích thước danh định cơ sở (Dài, Rộng, Cao) và tọa độ liên kết.  
2. **Khấu trừ tự động (Deduction Engine):** Hệ thống sử dụng danh sách openings và clashes được gán vào từng phần tử để trừ trực tiếp thể tích chiếm chỗ và diện tích ván khuôn tiếp xúc.  
3. **Không trùng lặp khối lượng (Zero Clashing Volume):** Sườn dầm đã trừ chiều dày sàn, đáy dầm đã trừ cột, đáy sàn đã trừ lưng dầm và diện tích cột/vách, đảm bảo bảng thống kê đạt chuẩn xác thực tế thi công.


### CÁC MODULE LOGIC BỔ SUNG CHO AI AGENT LẬP TRÌNH ADD-ON EXCEL

Dưới đây là mã giả và cấu trúc logic hoàn chỉnh cho các khoảng trống trên, được thiết kế theo đúng chuẩn đầu vào User Input -> Engine Calculation -> Output Table -> Link to BOQ Line Item.  
import math  
from dataclasses import dataclass, field  
from typing import List, Dict, Tuple  
from enum import Enum

\# ==============================================================================  
# MODULE BỔ SUNG 1: CÔNG TÁC CỌC CHUYÊN SÂU (ADVANCED PILING \& PRE-DRILLING)  
# Tương ứng các mã BOM: Factory F (1-12), Office (1-12), Bike Parking (1-10)...  
# ==============================================================================

@dataclass  
class PilingWorkItem:  
pile\_id: str  
pile\_dia\_mm: float             # 400 hoặc 600 mm  
total\_points: int              # Số lượng tim cọc (ea)  
depth\_per\_pile\_m: float        # Chiều sâu ép cọc thiết kế (m)  
ground\_level\_m: float          # Cốt đất tự nhiên hiện trạng (ví dụ +19.05m)  
cut\_off\_level\_m: float         # Cốt đỉnh cọc thiết kế (đáy đài) (ví dụ +18.55m)  
pre\_drill\_depth\_m: float = 0.0 # Chiều sâu khoan dẫn (thường 8 - 15m)  
need\_static\_load\_test: bool = True

&#x20;   def calculate\\\_piling\\\_boq(self) \\-\\> Dict\\\[str, float\\]:  
        """  
        Tự động tính: Cung cấp cọc, Khoan dẫn, Ép cọc âm, Cắt đầu cọc,   
        Thép bản đầu cọc, Nối cọc, Nén tĩnh.  
        """  
        \\# 1\\. Tổng chiều dài cọc cung cấp và ép (m)  
        total\\\_length\\\_m \\= self.total\\\_points \\\* self.depth\\\_per\\\_pile\\\_m  
          
        \\# 2\\. Chiều sâu ép cọc âm (m) \\= khoảng cách từ cốt đất tự nhiên xuống cốt đỉnh cọc  
        h\\\_negative\\\_m \\= max(0.0, self.ground\\\_level\\\_m \\- self.cut\\\_off\\\_level\\\_m)  
        total\\\_negative\\\_press\\\_m \\= self.total\\\_points \\\* h\\\_negative\\\_m  
          
        \\# 3\\. Số mối nối cọc hàn (giả định cọc chia đoạn tối đa 12m)  
        segments\\\_per\\\_pile \\= max(1, math.ceil(self.depth\\\_per\\\_pile\\\_m / 11.0))  
        welded\\\_joints\\\_count \\= self.total\\\_points \\\* (segments\\\_per\\\_pile \\- 1\\)  
          
        \\# 4\\. Thép bản bịt đầu cọc (Steel plate thk 2mm): 1 tấm/cọc  
        steel\\\_plate\\\_count \\= self.total\\\_points  
          
        \\# 5\\. Khoan dẫn tạo lỗ hạ cọc (ea)  
        pre\\\_drill\\\_count \\= self.total\\\_points if self.pre\\\_drill\\\_depth\\\_m \\> 0 else 0  
          
        \\# 6\\. Thí nghiệm nén tĩnh (Static Loading Test): Thường lấy 1% tổng số cọc (tối thiểu 2 cọc)  
        test\\\_piles \\= max(2, math.ceil(self.total\\\_points \\\* 0.01)) if self.need\\\_static\\\_load\\\_test else 0

        return {  
            "Total\\\_Pile\\\_Length\\\_m": round(total\\\_length\\\_m, 2),  
            "Pre\\\_Drilling\\\_Count\\\_ea": pre\\\_drill\\\_count,  
            "Negative\\\_Pressing\\\_Length\\\_m": round(total\\\_negative\\\_press\\\_m, 2),  
            "Cutting\\\_Pile\\\_Head\\\_Count\\\_ea": self.total\\\_points,  
            "Welded\\\_Joints\\\_ea": welded\\\_joints\\\_count,  
            "Steel\\\_Plate\\\_Head\\\_ea": steel\\\_plate\\\_count,  
            "Static\\\_Load\\\_Test\\\_ea": test\\\_piles,  
            "Transport\\\_Pile\\\_Length\\\_m": round(total\\\_length\\\_m, 2\\)  
        }


\# ==============================================================================  
# MODULE BỔ SUNG 2: SÀN NHÀ XƯỞNG CÔNG NGHIỆP (INDUSTRIAL GROUND SLAB \& FINISH)  
# Tương ứng mã BOM: Factory F (19-21, 35, 36, 77-87), Office (34, 35)  
# ==============================================================================

@dataclass  
class IndustrialGroundSlab:  
slab\_id: str  
length\_m: float                # Chiều dài khu vực sàn nhà xưởng (m)  
width\_m: float                 # Chiều rộng khu vực sàn (m)  
slab\_thickness\_mm: float = 200.0  # Chiều dày sàn BTCT nền (mm)  
base\_stone\_t\_mm: float = 250.0    # Đá cấp phối đầm chặt K98 (mm)  
joint\_grid\_spacing\_m: float = 6.0 # Lưới cắt khe co giãn sàn (6m x 6m)  
concrete\_grade: str = "M400"  
is\_steel\_fiber: bool = False      # Bê tông cốt sợi thép hay thép lưới thường  
is\_epoxy\_finish: bool = True      # Hoàn thiện sơn lót + tự san phẳng epoxy  
is\_hardener\_finish: bool = False  # Xoa tăng cứng Hardener

&#x20;   def calculate\\\_slab\\\_boq(self) \\-\\> Dict\\\[str, float\\]:  
        area\\\_m2 \\= self.length\\\_m \\\* self.width\\\_m  
        perimeter\\\_m \\= 2 \\\* (self.length\\\_m \\+ self.width\\\_m)  
          
        \\# 1\\. Đầm chặt đất hiện trạng K95 \& Trải màng PE 2 lớp (t=0.2mm)  
        subgrade\\\_compaction\\\_m2 \\= area\\\_m2  
        pe\\\_film\\\_m2 \\= area\\\_m2  
          
        \\# 2\\. Cấp phối đá dăm đầm chặt K98 (Hard cubic)  
        base\\\_stone\\\_m3 \\= area\\\_m2 \\\* (self.base\\\_stone\\\_t\\\_mm / 1000.0)  
          
        \\# 3\\. Bê tông sàn kết cấu (m3) \& Ván khuôn chặn cạnh sàn  
        concrete\\\_m3 \\= area\\\_m2 \\\* (self.slab\\\_thickness\\\_mm / 1000.0)  
        formwork\\\_edges\\\_m2 \\= perimeter\\\_m \\\* (self.slab\\\_thickness\\\_mm / 1000.0)  
          
        \\# 4\\. Đánh bóng và xoa nền (Power Float \& Polishing)  
        power\\\_float\\\_m2 \\= area\\\_m2  
          
        \\# 5\\. Cắt mạch bê tông sàn chống nứt (h=15mm, ô cờ 6m x 6m)  
        num\\\_cuts\\\_x \\= max(0, int(math.floor(self.length\\\_m / self.joint\\\_grid\\\_spacing\\\_m)))  
        num\\\_cuts\\\_y \\= max(0, int(math.floor(self.width\\\_m / self.joint\\\_grid\\\_spacing\\\_m)))  
        cutting\\\_joint\\\_m \\= (num\\\_cuts\\\_x \\\* self.width\\\_m) \\+ (num\\\_cuts\\\_y \\\* self.length\\\_m)  
          
        \\# 6\\. Sơn Epoxy hoàn thiện (Sơn lót \\+ Phủ 2 lớp hoặc Tự san phẳng 3mm)  
        epoxy\\\_area\\\_m2 \\= area\\\_m2 if self.is\\\_epoxy\\\_finish else 0.0  
        hardener\\\_area\\\_m2 \\= area\\\_m2 if self.is\\\_hardener\\\_finish else 0.0

        return {  
            "Subgrade\\\_Compaction\\\_m2": round(subgrade\\\_compaction\\\_m2, 2),  
            "PE\\\_Film\\\_2Plies\\\_m2": round(pe\\\_film\\\_m2, 2),  
            "Base\\\_Stone\\\_K98\\\_m3": round(base\\\_stone\\\_m3, 3),  
            "Slab\\\_Concrete\\\_m3": round(concrete\\\_m3, 3),  
            "Slab\\\_Formwork\\\_Edge\\\_m2": round(formwork\\\_edges\\\_m2, 2),  
            "Polishing\\\_Power\\\_Float\\\_m2": round(power\\\_float\\\_m2, 2),  
            "Joint\\\_Cutting\\\_m": round(cutting\\\_joint\\\_m, 2),  
            "Epoxy\\\_Coating\\\_m2": round(epoxy\\\_area\\\_m2, 2),  
            "Hardener\\\_Floor\\\_m2": round(hardener\\\_area\\\_m2, 2\\)  
        }


\# ==============================================================================  
# MODULE BỔ SUNG 3: HỆ THỐNG MÁI NHÀ XƯỞNG \& TƯỜNG CHẮN PARAPET (ROOFING SYSTEM)  
# Tương ứng mã BOM: Factory F (1-31), Office (1-15), Bike Parking (1-11)  
# ==============================================================================

@dataclass  
class IndustrialRoofSystem:  
roof\_id: str  
length\_m: float                # Chiều dài mái bê tông (m)  
width\_m: float                 # Chiều rộng mái bê tông (m)  
parapet\_height\_m: float = 1.0  # Chiều cao tường chắn mái Parapet (m)  
parapet\_thickness\_mm: float = 200.0 # Chiều dày tường Parapet (mm)  
screed\_slope\_avg\_t\_mm: float = 100.0 # Bê tông đổ bù tạo dốc (M250, dày tb 100mm)  
roof\_drain\_spacing\_m: float = 18.0   # Khoảng cách bố trí phễu thu nước mái (m)  
downspout\_height\_m: float = 15.0     # Chiều cao ống thoát nước từ mái xuống đất (m)

&#x20;   def calculate\\\_roof\\\_boq(self) \\-\\> Dict\\\[str, float\\]:  
        area\\\_roof\\\_m2 \\= self.length\\\_m \\\* self.width\\\_m  
        perimeter\\\_parapet\\\_m \\= 2 \\\* (self.length\\\_m \\+ self.width\\\_m)  
          
        \\# 1\\. Bê tông đổ bù tạo dốc mác M250 (dày trung bình 100mm) \\+ Lưới thép hàn D8 ô 150x150  
        slope\\\_concrete\\\_m3 \\= area\\\_roof\\\_m2 \\\* (self.screed\\\_slope\\\_avg\\\_t\\\_mm / 1000.0)  
        wire\\\_mesh\\\_m2 \\= area\\\_roof\\\_m2  
          
        \\# 2\\. Hệ chống thấm Urethane đa lớp theo spec BOM:  
        \\# \\- Lớp chống thấm không lộ thiên dầy 3mm (Non-exposed Urethane)  
        \\# \\- Lớp phủ Urethane lộ thiên kháng tia cực tím dầy 2mm (Exposed Urethane UV protection)  
        urethane\\\_non\\\_exposed\\\_m2 \\= area\\\_roof\\\_m2  
        urethane\\\_exposed\\\_m2 \\= area\\\_roof\\\_m2  
        power\\\_float\\\_roof\\\_m2 \\= area\\\_roof\\\_m2  
          
        \\# Khe co giãn trên mái (cắt và bơm keo PU sealant)  
        expansion\\\_joint\\\_roof\\\_m \\= (self.length\\\_m / 6.0 \\\* self.width\\\_m) \\+ (self.width\\\_m / 6.0 \\\* self.length\\\_m)  
          
        \\# 3\\. Tường chắn Parapet hoàn thiện:  
        \\# \\- Khối xây gạch Parapet: Chu vi x Chiều cao x Chiều dày  
        v\\\_brick\\\_parapet\\\_m3 \\= (perimeter\\\_parapet\\\_m \\\* self.parapet\\\_height\\\_m \\\*   
                              (self.parapet\\\_thickness\\\_mm / 1000.0))  
        \\# \\- Trát ngoài và quét chống thấm Urethane thành Parapet (mặt trong và đỉnh):  
        plaster\\\_parapet\\\_m2 \\= perimeter\\\_parapet\\\_m \\\* (self.parapet\\\_height\\\_m \\+ (self.parapet\\\_thickness\\\_mm / 1000.0))  
        urethane\\\_parapet\\\_m2 \\= plaster\\\_parapet\\\_m2  
          
        \\# 4\\. Hệ thống máng xối và ống thu nước mái:  
        \\# Số lượng phễu thu gang kèm chắn rác đúc sẵn  
        num\\\_roof\\\_drains \\= max(4, int(math.ceil(perimeter\\\_parapet\\\_m / self.roof\\\_drain\\\_spacing\\\_m)))  
        \\# Ống thoát nước mái SUS304 DN150/DN200  
        total\\\_downspout\\\_pipe\\\_m \\= num\\\_roof\\\_drains \\\* self.downspout\\\_height\\\_m  
        pipe\\\_brackets\\\_count \\= num\\\_roof\\\_drains \\\* int(math.ceil(self.downspout\\\_height\\\_m / 1.5)) \\# Giá đỡ 1.5m/cái  
        elbow\\\_90\\\_count \\= num\\\_roof\\\_drains \\\* 2 \\# Cút 90 độ chuyển hướng

        return {  
            "Screed\\\_Slope\\\_Concrete\\\_M250\\\_m3": round(slope\\\_concrete\\\_m3, 3),  
            "Roof\\\_Wire\\\_Mesh\\\_D8\\\_m2": round(wire\\\_mesh\\\_m2, 2),  
            "Urethane\\\_Non\\\_Exposed\\\_3mm\\\_m2": round(urethane\\\_non\\\_exposed\\\_m2, 2),  
            "Urethane\\\_Exposed\\\_UV\\\_2mm\\\_m2": round(urethane\\\_exposed\\\_m2, 2),  
            "Roof\\\_Power\\\_Float\\\_m2": round(power\\\_float\\\_roof\\\_m2, 2),  
            "Roof\\\_Expansion\\\_Joint\\\_m": round(expansion\\\_joint\\\_roof\\\_m, 2),  
            "Parapet\\\_Brick\\\_m3": round(v\\\_brick\\\_parapet\\\_m3, 3),  
            "Parapet\\\_Plaster\\\_m2": round(plaster\\\_parapet\\\_m2, 2),  
            "Parapet\\\_Urethane\\\_m2": round(urethane\\\_parapet\\\_m2, 2),  
            "Roof\\\_Drain\\\_Cast\\\_Iron\\\_ea": num\\\_roof\\\_drains,  
            "SUS304\\\_Downspout\\\_Pipe\\\_m": round(total\\\_downspout\\\_pipe\\\_m, 2),  
            "Downspout\\\_Brackets\\\_ea": pipe\\\_brackets\\\_count,  
            "Downspout\\\_Elbows\\\_ea": elbow\\\_90\\\_count  
        }


\# ==============================================================================  
# MODULE BỔ SUNG 4: VÁCH NGĂN PANEL \& HỆ THỐNG CỬA (PANEL \& DOOR SCHEDULE)  
# Tương ứng mã BOM: Factory F (35-41, 69-76, 127-143), Office (43-45, 74-87)  
# ==============================================================================

@dataclass  
class IndustrialDoorSpec:  
door\_code: str               # 1FSD1, 2FSD1, 1SD1, 1ST (Shutter), 1PW...  
width\_m: float  
height\_m: float  
quantity: int  
door\_type: str               # "FIRE\_STEEL", "STEEL", "SHUTTER", "ABS", "ALU\_WINDOW"  
fire\_rating: str = "EI60"

&#x20;   @property  
    def total\\\_area\\\_m2(self) \\-\\> float:  
        return self.width\\\_m \\\* self.height\\\_m \\\* self.quantity


@dataclass  
class IndustrialPanelPartition:  
partition\_id: str  
length\_m: float              # Chiều dài tuyến vách (m)  
height\_m: float              # Chiều cao vách panel (m)  
panel\_thickness\_mm: float = 50.0 # 50mm, 75mm hoặc 100mm PU  
doors: List\[IndustrialDoorSpec] = field(default\_factory=list)

&#x20;   def calculate\\\_panel\\\_boq(self) \\-\\> Dict\\\[str, float\\]:  
        gross\\\_area\\\_m2 \\= self.length\\\_m \\\* self.height\\\_m  
        door\\\_cut\\\_area\\\_m2 \\= sum(d.total\\\_area\\\_m2 for d in self.doors)  
        net\\\_panel\\\_area\\\_m2 \\= max(0.0, gross\\\_area\\\_m2 \\- door\\\_cut\\\_area\\\_m2)  
          
        \\# 1\\. Phụ kiện nẹp nhôm định hình:  
        \\# \\- Thanh nẹp U nhôm chân vách (Base U-channel): Dọc chiều dài sàn trừ cửa đi  
        door\\\_width\\\_sum \\= sum(d.width\\\_m \\\* d.quantity for d in self.doors if "WINDOW" not in d.door\\\_type)  
        base\\\_u\\\_channel\\\_m \\= max(0.0, self.length\\\_m \\- door\\\_width\\\_sum)  
          
        \\# \\- Thanh nẹp V nhôm ốp đỉnh (Top cover) và nẹp viền cửa  
        top\\\_v\\\_cover\\\_m \\= self.length\\\_m  
        jamb\\\_flashing\\\_m \\= sum((2 \\\* d.height\\\_m \\+ d.width\\\_m) \\\* d.quantity for d in self.doors)  
          
        \\# 2\\. Khung thép hộp đỡ vách panel (Square pipe 100x100x3 hoặc 50x50x2.3):  
        \\# Trung bình định mức 8 \\- 10 kg thép hộp / m2 vách panel  
        steel\\\_support\\\_frame\\\_kg \\= net\\\_panel\\\_area\\\_m2 \\\* 8.5

        return {  
            "Gross\\\_Panel\\\_Area\\\_m2": round(gross\\\_area\\\_m2, 2),  
            "Door\\\_Deduction\\\_Area\\\_m2": round(door\\\_cut\\\_area\\\_m2, 2),  
            "Net\\\_Panel\\\_Area\\\_m2": round(net\\\_panel\\\_area\\\_m2, 2),  
            "Base\\\_U\\\_Aluminum\\\_m": round(base\\\_u\\\_channel\\\_m, 2),  
            "Top\\\_V\\\_Aluminum\\\_m": round(top\\\_v\\\_cover\\\_m, 2),  
            "Door\\\_Flashing\\\_m": round(jamb\\\_flashing\\\_m, 2),  
            "Steel\\\_Support\\\_Frame\\\_kg": round(steel\\\_support\\\_frame\\\_kg, 2\\)  
        }


\# ==============================================================================  
# MODULE BỔ SUNG 5: HẠ TẦNG ĐƯỜNG NHỰA ASPHALT \& VỈA HÈ (ASPHALT ROAD \& PAVEMENT)  
# Tương ứng mã BOM: Internal Road (4-5), Pavement finishing (6-10), Grease Trap  
# ==============================================================================

@dataclass  
class AsphaltRoadAndSidewalk:  
road\_id: str  
road\_length\_m: float  
road\_width\_m: float  
sidewalk\_width\_m: float = 1.5   # Bề rộng vỉa hè mỗi bên  
asphalt\_thickness\_cm: float = 5.0 # Bê tông nhựa hạt mịn C12.5 dày 5cm  
tack\_coat\_rate\_kg\_m2: float = 0.5 # Tưới nhũ tương bám dính 0.5 kg/m2

&#x20;   def calculate\\\_asphalt\\\_pavement\\\_boq(self) \\-\\> Dict\\\[str, float\\]:  
        road\\\_area\\\_m2 \\= self.road\\\_length\\\_m \\\* self.road\\\_width\\\_m  
        sidewalk\\\_area\\\_m2 \\= 2 \\\* (self.road\\\_length\\\_m \\\* self.sidewalk\\\_width\\\_m)  
          
        \\# 1\\. Bê tông nhựa hạt mịn C12.5 dày 5cm (m2)  
        asphalt\\\_concrete\\\_c12\\\_5\\\_m2 \\= road\\\_area\\\_m2  
          
        \\# 2\\. Tưới nhũ tương bám dính Bituminous Tack Coat 0.5 kg/m2 (m2)  
        tack\\\_coat\\\_m2 \\= road\\\_area\\\_m2  
          
        \\# 3\\. Lối đi bộ vỉa hè: Đầm đất K95, Lót bê tông M150 dày 80mm, Láng vữa M75 dày 20mm, Lát gạch Terrazzo 400x400x30  
        sidewalk\\\_compaction\\\_m2 \\= sidewalk\\\_area\\\_m2  
        sidewalk\\\_concrete\\\_m150\\\_m3 \\= sidewalk\\\_area\\\_m2 \\\* 0.08  
        sidewalk\\\_mortar\\\_bed\\\_m2 \\= sidewalk\\\_area\\\_m2  
        terrazzo\\\_tiles\\\_m2 \\= sidewalk\\\_area\\\_m2  
          
        \\# 4\\. Bể tách mỡ nhà ăn (Grease Trap đúc sẵn/xây gạch trát vữa M75): 1 bộ chuẩn xưởng  
        grease\\\_trap\\\_sets \\= 1

        return {  
            "Asphalt\\\_Concrete\\\_C12\\\_5\\\_m2": round(asphalt\\\_concrete\\\_c12\\\_5\\\_m2, 2),  
            "Tack\\\_Coat\\\_Emulsion\\\_m2": round(tack\\\_coat\\\_m2, 2),  
            "Sidewalk\\\_Terrazzo\\\_Tiles\\\_m2": round(terrazzo\\\_tiles\\\_m2, 2),  
            "Sidewalk\\\_Concrete\\\_M150\\\_m3": round(sidewalk\\\_concrete\\\_m150\\\_m3, 3),  
            "Sidewalk\\\_Mortar\\\_Bed\\\_m2": round(sidewalk\\\_mortar\\\_bed\\\_m2, 2),  
            "Grease\\\_Trap\\\_Sets": grease\\\_trap\\\_sets  
        }


### KIẾN TRÚC TÍCH HỢP CHO AI AGENT VIẾT ADD-ON EXCEL

Để AI Agent lập trình ra Add-on Excel tự động hóa từ đầu vào đến bảng BOQ hoàn chỉnh, bạn chỉ cần hướng dẫn Agent triển khai theo mô hình **3 Tầng (3-Tier Architecture)**:  
\[Sheet 1: USER\_INPUT\_FORM] (Giao diện nhập kích thước hình học)  
│  - Nhập tên cấu kiện (Cột, Dầm, Sàn, Mái, Cống...)  
│  - Nhập Kích thước danh định: Dài (L), Rộng (B), Cao (H), Chiều dày (T)  
│  - Nhập danh sách lỗ mở (Cửa, Lỗ MEP) và liên kết cấu kiện  
▼  
\[ADDON CORE: SPEC \& CALCULATION ENGINE] (Các module Python/VBA đã cung cấp)  
│  - Áp dụng trừ giao cắt tự động (Cột > Dầm > Sàn)  
│  - Tự động tra bảng Chiều dài neo Lan, Nối chồng Llap, Độ giãn góc uốn  
│  - Tự động sinh danh mục thanh thép (Bar Bending Schedule - BBS)  
│  - Tự động tính các lớp bồi hoàn thiện, máng xối, nẹp nhôm, vỉa hè  
▼  
\[Sheet 2: CALCULATION\_SHEET] (Bảng tính chi tiết diễn giải từng cấu kiện)  
│  - Bảng thống kê thép chi tiết (Số hiệu, Hình dạng, Chiều dài, Số lượng, Kg)  
│  - Bảng diễn giải chi tiết khối lượng Bê tông, Ván khuôn, Đào đất, Hoàn thiện  
▼  
\[Sheet 3: MASTER\_BOQ] (Tự động Link tổng khối lượng theo đúng đầu mục của file BOM)  
1. PHC Pile Works (m cọc, đầu cọc, mối nối)  
2. Soil Works (m3 đào, m3 đắp K95, m3 cát đệm)  
3. Concrete \& Formwork Works (m3 bê tông M400, m2 ván khuôn)  
4. Rebar Works (Tấn thép CB500-V, CB240-T)  
5. Finishing Works (m2 trát, m2 sơn, m2 Urethane, m2 panel, bộ cửa)  
6. Infrastructure Works (m2 đường bê tông/nhựa, m cống, hố ga)  
Toàn bộ hệ thống quy tắc và mã logic đã bao phủ đầy đủ và khớp với danh mục thi công thực tế của dự án nhà xưởng trong **BOM.pdf**, sẵn sàng để chuyển giao cho AI Agent bắt tay vào viết code Add-on.


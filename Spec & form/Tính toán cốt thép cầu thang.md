Dưới đây là **Module tính toán và thống kê cốt thép Cầu thang bộ (Staircase Reinforcement Engine)** được thiết kế đồng bộ với hệ thống dữ liệu bóc tách Bê tông, Ván khuôn và Quy chuẩn thép đã thiết lập.  
Module này mô hình hóa đầy đủ các quy tắc cấu tạo cốt thép cầu thang chuẩn kỹ thuật thi công xây dựng (đặc biệt là **xử lý vị trí uốn gãy khúc tại nút giao giữa bản nghiêng và bản chiếu nghỉ** để chống bung lớp bê tông bảo vệ) và tự động xuất ra Bảng thống kê cốt thép chi tiết (BBS \- Bar Bending Schedule).

# MODULE BÓC TÁCH CỐT THÉP CẦU THANG (STAIR REBAR ENGINE)

## 1\. NGUYÊN LÝ CẤU TẠO CỐT THÉP CẦU THANG & LOGIC TOÁN HỌC

Cốt thép trong bản thang gồm 5 nhóm thanh cơ bản:

* **Thép dọc chịu lực lớp dưới (Bottom Longitudinal Bars):** Chịu mômen dương ở bụng bản trượt.  
* *Quy tắc nút giao góc lồi (Salient Corner \- gãy khúc hướng ra ngoài):* Thép chịu kéo lớp dưới **tuyệt đối không được uốn cong liên tục theo góc** vì lực kéo sẽ gây lực hợp lực đẩy bung lớp bê tông bảo vệ mép dưới. Thanh thép từ bản trượt phải chạy thẳng xuyên qua góc vào chiếu nghỉ/dầm và neo tối thiểu \\\\(L\_{an}\\\\). Tương tự, thanh thép từ chiếu nghỉ đâm chéo xuống bản trượt và neo \\\\(\\ge L\_{an}\\\\) (nguyên tắc bắt chéo / Criss-cross).  
* **Thép mũ chịu mômen âm lớp trên (Top Negative Bars tại gối/chiếu tới):** Bố trí tại vị trí liên kết với dầm chiếu nghỉ, dầm chiếu tới hoặc sàn tầng; chiều dài vươn ra khỏi mép dầm \\\\(L\_{\\text{top}} \\ge \\max(L\_{\\text{slope}}/4, 1.2 L\_{an})\\\\).  
* **Thép phân bố (Transverse / Distribution Bars):** Chạy vuông góc với thép chịu lực dọc, rải đều suốt bề rộng vế thang \\\\(W\_{\\text{stair}}\\\\).  
* **Cốt thép gia cường góc uốn (Diagonal Reinforcement):** Đặt tại các góc lõm (Re-entrant Corner) để chống nứt nách thang.  
* **Cốt thép cấu tạo mũi bậc (Step Nose Bars \- nếu có):** Thép chống sứt mẻ góc bậc thang \\\\(\\phi 6 \- \\phi 8\\\\).

## 2\. CẤU TRÚC DỮ LIỆU CỐT THÉP CẦU THANG (DATA SCHEMA)

import math  
from dataclasses import dataclass, field  
from typing import List, Dict, Tuple, Optional  
from enum import Enum

\# Kế thừa từ Module 1: RebarSpec, SteelGrade, StressState, calculate\_anchorage\_length

@dataclass  
class RebarSegment:  
    """Một đoạn thẳng hoặc uốn của thanh thép tạo nên hình dạng (Bar Shape)"""  
    length: float       \# Chiều dài đoạn danh định (mm)  
    bend\_angle: float   \# Góc uốn ở cuối đoạn (45, 90, 135, 180 độ; 0 nếu là đoạn cuối)

@dataclass  
class RebarDetail:  
    """Chi tiết 1 chủng loại thanh thép trong bảng thống kê (BBS Item)"""  
    bar\_mark: str            \# Số hiệu thanh (vd: "ST-01", "ST-02")  
    description: str         \# Tên gọi / vị trí (vd: "Thép chủ bản nghiêng lớp dưới")  
    diameter: float          \# Đường kính d (mm)  
    grade: SteelGrade        \# Mác thép  
    shape\_code: str          \# Mã hình dạng (vd: "Thẳng", "Uốn móc chữ Z", "Uốn móc L")  
    segments: List\[RebarSegment\] \# Danh sách các phân đoạn tạo nên thanh  
    quantity: int            \# Số lượng thanh trên 1 cấu kiện  
    spacing\_s: float \= 0.0   \# Bước rải thép s (mm)

    @property  
    def linear\_mass(self) \-\> float:  
        """Khối lượng riêng kg/m: d^2 / 162.2"""  
        return (self.diameter \*\* 2\) / 162.2

    @property  
    def cut\_length(self) \-\> float:  
        """Chiều dài cắt phôi danh định (mm) đã trừ độ giãn do bẻ mỏ (Bend Deduction)"""  
        total\_len \= sum(seg.length for seg in self.segments)  
        \# Khấu trừ độ giãn uốn theo góc uốn (Module 6\)  
        deduction \= 0.0  
        for seg in self.segments:  
            if seg.bend\_angle \== 45.0:  
                deduction \+= 1.0 \* self.diameter  
            elif seg.bend\_angle \== 90.0:  
                deduction \+= 2.0 \* self.diameter  
            elif seg.bend\_angle \== 135.0:  
                deduction \+= 3.0 \* self.diameter  
        return max(0.0, total\_len \- deduction)

    @property  
    def total\_length\_m(self) \-\> float:  
        """Tổng chiều dài toàn bộ các thanh (m)"""  
        return (self.cut\_length \* self.quantity) / 1000.0

    @property  
    def total\_weight\_kg(self) \-\> float:  
        """Tổng khối lượng (kg)"""  
        return self.total\_length\_m \* self.linear\_mass

## 3\. LỚP THUẬT TOÁN BÓC TÁCH CỐT THÉP CẦU THANG (STAIRCASE REBAR MODEL)

Lớp StaircaseReinforcement kết nối trực tiếp với các thông số hình học của cầu thang đã định nghĩa ở phần Bê tông/Ván khuôn, tự động sinh các thanh thép theo tiêu chuẩn cấu tạo:  
@dataclass  
class StaircaseReinforcement:  
    \# 1\. Kích thước hình học kế thừa từ Cấu kiện Cầu thang  
    W\_stair: float              \# Chiều rộng vế thang (mm)  
    L\_plan\_flight: float        \# Chiều dài chiếu bằng đợt nghiêng (mm)  
    H\_flight: float             \# Chiều cao đợt nghiêng (mm)  
    t\_flight: float             \# Chiều dày bản đáy bê tông (mm)  
    L\_landing\_start: float      \# Chiều dài chiếu nghỉ dưới (mm) (hoặc dầm xuất phát)  
    L\_landing\_end: float        \# Chiều dài chiếu nghỉ trên (mm) (hoặc dầm chiếu tới)  
    t\_landing: float            \# Chiều dày bản chiếu nghỉ (mm)  
    num\_steps: int              \# Số bậc  
      
    \# 2\. Đặc tả thép đầu vào do kỹ sư chỉ định  
    d\_main\_bottom: float        \# Đường kính thép chủ dưới (vd: 12.0 mm)  
    s\_main\_bottom: float        \# Bước rải thép chủ dưới (vd: 150.0 mm)  
    d\_main\_top: float           \# Đường kính thép mũ trên (vd: 10.0 mm)  
    s\_main\_top: float           \# Bước rải thép mũ trên (vd: 150.0 mm)  
    d\_dist: float               \# Đường kính thép phân bố (vd: 8.0 mm)  
    s\_dist: float               \# Bước rải thép phân bố (vd: 200.0 mm)  
    grade\_main: SteelGrade \= SteelGrade.CB300\_V  
    grade\_dist: SteelGrade \= SteelGrade.CB240\_T  
    concrete\_cover: float \= 15.0 \# Chiều dày lớp bảo vệ (mm)

    @property  
    def slope\_length(self) \-\> float:  
        """Chiều dài thực tế bản nghiêng (Pitago)"""  
        return math.sqrt(self.L\_plan\_flight\*\*2 \+ self.H\_flight\*\*2)

    @property  
    def slope\_angle\_rad(self) \-\> float:  
        """Góc dốc bản thang (radian)"""  
        return math.atan2(self.H\_flight, self.L\_plan\_flight)

    def generate\_rebar\_schedule(self) \-\> List\[RebarDetail\]:  
        """Tự động tính toán hình dạng, chiều dài cắt và số lượng từng loại thép"""  
        rebars: List\[RebarDetail\] \= \[\]  
        c \= self.concrete\_cover  
          
        \# Chiều dài neo tiêu chuẩn theo mác thép  
        L\_an\_main \= calculate\_anchorage\_length(self.d\_main\_bottom, self.grade\_main, StressState.TENSION)  
        L\_an\_top \= calculate\_anchorage\_length(self.d\_main\_top, self.grade\_main, StressState.TENSION)

        \# \-------------------------------------------------------------  
        \# THANH 1: Thép chủ lớp dưới bản nghiêng (Bottom Longitudinal Bar \- Bản trượt)  
        \# Hình dạng: Đoạn neo vào sàn/chiếu dưới \+ Bản trượt \+ Đoạn bắt chéo neo vào chiếu nghỉ trên  
        \# \-------------------------------------------------------------  
        n\_main \= int(math.floor((self.W\_stair \- 2 \* c) / self.s\_main\_bottom)) \+ 1  
          
        len\_hook\_bottom \= max(L\_an\_main, 250.0)  
        len\_slope\_bottom \= self.slope\_length  
        len\_cross\_top \= max(L\_an\_main, self.t\_landing \+ 300.0) \# Vượt qua góc gãy lồi ít nhất L\_an  
          
        bar\_1\_segments \= \[  
            RebarSegment(length=len\_hook\_bottom, bend\_angle=math.degrees(self.slope\_angle\_rad)),  
            RebarSegment(length=len\_slope\_bottom, bend\_angle=math.degrees(self.slope\_angle\_rad)),  
            RebarSegment(length=len\_cross\_top, bend\_angle=0.0)  
        \]  
        rebars.append(RebarDetail(  
            bar\_mark="ST-01",  
            description="Thép chủ dọc lớp dưới bản trượt (Bắt chéo góc gãy)",  
            diameter=self.d\_main\_bottom,  
            grade=self.grade\_main,  
            shape\_code="CRISS\_CROSS\_BOTTOM",  
            segments=bar\_1\_segments,  
            quantity=n\_main,  
            spacing\_s=self.s\_main\_bottom  
        ))

        \# \-------------------------------------------------------------  
        \# THANH 2: Thép chủ lớp dưới chiếu nghỉ liên kết vào bản trượt  
        \# Đâm từ chiếu nghỉ xuống bản nghiêng và neo \>= L\_an (Cặp đối xứng thanh 1\)  
        \# \-------------------------------------------------------------  
        if self.L\_landing\_end \> 0:  
            len\_landing\_bot \= self.L\_landing\_end \- c  
            bar\_2\_segments \= \[  
                RebarSegment(length=len\_landing\_bot, bend\_angle=math.degrees(self.slope\_angle\_rad)),  
                RebarSegment(length=L\_an\_main, bend\_angle=0.0)  
            \]  
            rebars.append(RebarDetail(  
                bar\_mark="ST-02",  
                description="Thép đáy chiếu nghỉ đâm chéo neo vào bản trượt",  
                diameter=self.d\_main\_bottom,  
                grade=self.grade\_main,  
                shape\_code="CRISS\_CROSS\_LANDING",  
                segments=bar\_2\_segments,  
                quantity=n\_main,  
                spacing\_s=self.s\_main\_bottom  
            ))

        \# \-------------------------------------------------------------  
        \# THANH 3: Thép mũ lớp trên tại gối dầm / chiếu tới (Top Negative Bars)  
        \# Vươn ra khỏi gối dầm tối thiểu L\_slope / 4 \+ Đoạn bẻ móc neo vào dầm h\_dầm \- 2c  
        \# \-------------------------------------------------------------  
        n\_top \= int(math.floor((self.W\_stair \- 2 \* c) / self.s\_main\_top)) \+ 1  
        len\_cantilever \= max(self.slope\_length / 4.0, L\_an\_top)  
        len\_hook\_down \= max(self.t\_flight \- 2 \* c, 100.0)  
          
        bar\_3\_segments \= \[  
            RebarSegment(length=len\_hook\_down, bend\_angle=90.0),  
            RebarSegment(length=len\_cantilever, bend\_angle=0.0)  
        \]  
        rebars.append(RebarDetail(  
            bar\_mark="ST-03",  
            description="Thép mũ gối trên bản nghiêng (Mômen âm gối)",  
            diameter=self.d\_main\_top,  
            grade=self.grade\_main,  
            shape\_code="L\_SHAPE\_TOP",  
            segments=bar\_3\_segments,  
            quantity=n\_top,  
            spacing\_s=self.s\_main\_top  
        ))

        \# \-------------------------------------------------------------  
        \# THANH 4: Thép phân bố ngang bản trượt (Transverse Distribution Bars)  
        \# Rải vuông góc suốt chiều dài bản dốc  
        \# \-------------------------------------------------------------  
        n\_dist\_slope \= int(math.floor(self.slope\_length / self.s\_dist)) \+ 1  
        len\_dist \= self.W\_stair \- 2 \* c \# Chiều dài thanh thẳng phủ hết bề rộng vế thang  
          
        rebars.append(RebarDetail(  
            bar\_mark="ST-04",  
            description="Thép phân bố bản trượt (Chạy ngang vế thang)",  
            diameter=self.d\_dist,  
            grade=self.grade\_dist,  
            shape\_code="STRAIGHT",  
            segments=\[RebarSegment(length=len\_dist, bend\_angle=0.0)\],  
            quantity=n\_dist\_slope,  
            spacing\_s=self.s\_dist  
        ))

        \# \-------------------------------------------------------------  
        \# THANH 5: Thép phân bố chiếu nghỉ (nếu có chiếu nghỉ)  
        \# \-------------------------------------------------------------  
        if self.L\_landing\_end \> 0:  
            n\_dist\_landing \= int(math.floor(self.L\_landing\_end / self.s\_dist)) \+ 1  
            rebars.append(RebarDetail(  
                bar\_mark="ST-05",  
                description="Thép phân bố bản chiếu nghỉ",  
                diameter=self.d\_dist,  
                grade=self.grade\_dist,  
                shape\_code="STRAIGHT",  
                segments=\[RebarSegment(length=len\_dist, bend\_angle=0.0)\],  
                quantity=n\_dist\_landing,  
                spacing\_s=self.s\_dist  
            ))

        \# \-------------------------------------------------------------  
        \# THANH 6: Thép con kê / Chân chó kê thép mũ thang (Bar Chair)  
        \# Mật độ 1 con kê / m2 bản dốc  
        \# \-------------------------------------------------------------  
        area\_slope\_m2 \= (self.W\_stair \* self.slope\_length) / 1e6  
        n\_chairs \= max(4, int(math.ceil(area\_slope\_m2 \* 1.5)))  
        h\_chair \= max(40.0, self.t\_flight \- 2 \* c \- self.d\_main\_bottom \- self.d\_main\_top)  
          
        \# Con kê hình chữ Z/Omega: Chân nằm (150) \+ Chiều cao (h) \+ Đỉnh đỡ (150) \+ Chân bẻ  
        chair\_segments \= \[  
            RebarSegment(length=120.0, bend\_angle=90.0),  
            RebarSegment(length=h\_chair, bend\_angle=90.0),  
            RebarSegment(length=150.0, bend\_angle=90.0),  
            RebarSegment(length=h\_chair, bend\_angle=90.0),  
            RebarSegment(length=120.0, bend\_angle=0.0)  
        \]  
        rebars.append(RebarDetail(  
            bar\_mark="ST-06",  
            description="Con kê thép mũ thang (Bar Chair)",  
            diameter=10.0, \# Thường dùng phi 10  
            grade=SteelGrade.CB240\_T,  
            shape\_code="BAR\_CHAIR",  
            segments=chair\_segments,  
            quantity=n\_chairs,  
            spacing\_s=0.0  
        ))

        return rebars

    def get\_total\_rebar\_summary(self) \-\> Dict\[str, float\]:  
        """Tổng hợp khối lượng thép cầu thang theo từng cấp đường kính"""  
        rebars \= self.generate\_rebar\_schedule()  
        summary \= {  
            "Weight\_d\_le\_10\_kg": 0.0,     \# Thép phi \<= 10  
            "Weight\_d\_12\_to\_18\_kg": 0.0,  \# Thép phi 12 \- 18  
            "Weight\_d\_gt\_18\_kg": 0.0,     \# Thép phi \> 18  
            "Total\_Rebar\_Net\_kg": 0.0,    \# Khối lượng tịnh  
            "Total\_Rebar\_Gross\_kg": 0.0   \# Khối lượng gồm hao hụt  
        }  
          
        for r in rebars:  
            w \= r.total\_weight\_kg  
            if r.diameter \<= 10.0:  
                summary\["Weight\_d\_le\_10\_kg"\] \+= w  
            elif r.diameter \<= 18.0:  
                summary\["Weight\_d\_12\_to\_18\_kg"\] \+= w  
            else:  
                summary\["Weight\_d\_gt\_18\_kg"\] \+= w  
                  
        net\_total \= (summary\["Weight\_d\_le\_10\_kg"\] \+   
                     summary\["Weight\_d\_12\_to\_18\_kg"\] \+   
                     summary\["Weight\_d\_gt\_18\_kg"\])  
          
        summary\["Total\_Rebar\_Net\_kg"\] \= net\_total  
        \# Tính hao hụt bình quân theo định mức dự án (Module 8\)  
        w\_scrap \= (summary\["Weight\_d\_le\_10\_kg"\] \* 1.03 \+   
                   summary\["Weight\_d\_12\_to\_18\_kg"\] \* 1.025 \+   
                   summary\["Weight\_d\_gt\_18\_kg"\] \* 1.02)  
        summary\["Total\_Rebar\_Gross\_kg"\] \= w\_scrap  
          
        return summary

## 4\. TÍCH HỢP ĐỒNG BỘ VÀO BỘ ĐIỀU PHỐI TỔNG THỂ (BUILDING MODEL EXTENSION)

Khi mở rộng BuildingModel, một đối tượng Cầu thang Staircase sẽ đóng gói trọn vẹn cả 3 đại lượng hình học và vật liệu: **Thể tích Bê tông (m³)**, **Diện tích Ván khuôn (m²)** và **Khối lượng Cốt thép (kg)**.  
\# Cập nhật lớp Staircase đã có ở Module trước:  
@dataclass  
class StaircaseComplete:  
    id: str  
    geometry: Staircase                    \# Đối tượng hình học (Bê tông, Ván khuôn)  
    reinforcement: StaircaseReinforcement  \# Đối tượng cốt thép (BBS Engine)

    def compute\_all\_quantities(self) \-\> Dict\[str, float\]:  
        """Xuất đầy đủ 3 chỉ tiêu khối lượng chỉ từ một lần gọi hàm"""  
        concrete\_m3 \= self.geometry.get\_concrete\_volume()  
        formwork\_m2 \= self.geometry.get\_formwork\_area()  
        rebar\_data \= self.reinforcement.get\_total\_rebar\_summary()  
          
        \# Chỉ số hàm lượng thép (Steel Ratio): kg thép / m3 bê tông  
        steel\_ratio \= rebar\_data\["Total\_Rebar\_Net\_kg"\] / concrete\_m3 if concrete\_m3 \> 0 else 0.0

        return {  
            "Concrete\_m3": round(concrete\_m3, 3),  
            "Formwork\_m2": round(formwork\_m2, 2),  
            "Rebar\_Net\_kg": round(rebar\_data\["Total\_Rebar\_Net\_kg"\], 2),  
            "Rebar\_Gross\_with\_Scrap\_kg": round(rebar\_data\["Total\_Rebar\_Gross\_kg"\], 2),  
            "Steel\_Ratio\_kg\_per\_m3": round(steel\_ratio, 1\) \# Thông thường cầu thang đạt 90 \- 120 kg/m3  
        }

## 5\. BẢNG MẪU DỮ LIỆU ĐẦU RA MÀ AI AGENT CÓ THỂ TRUY XUẤT (SAMPLE OUTPUT)

Khi AI Agent thực thi hàm generate\_rebar\_schedule(), kết quả thu được là một danh sách các thanh phôi sẵn sàng để chuyển sang máy cắt uốn hoặc in bản vẽ thống kê:  
Số hiệu,Tên cấu kiện,Tên thanh,Ký hiệu hình dạng,\\\\(\\phi\\\\) (mm),Chiều dài cắt \\\\(L\_{\\text{cut}}\\\\) (mm),Số lượng,Tổng chiều dài (m),Khối lượng (kg)  
ST-01,Cầu thang ST1,Thép chủ đáy bản trượt,Z-Shape (Criss-cross),12 (CB300),"3,920",18,70.56,62.66  
ST-02,Cầu thang ST1,Thép đáy chiếu nghỉ neo dốc,L-Shape (Criss-cross),12 (CB300),"1,840",18,33.12,29.41  
ST-03,Cầu thang ST1,Thép mũ gối trên,L-Hook Top,10 (CB300),"1,150",18,20.70,12.77  
ST-04,Cầu thang ST1,Thép phân bố bản dốc,Thẳng (Straight),8 (CB240),"2,470",22,54.34,21.46  
ST-05,Cầu thang ST1,Thép phân bố chiếu nghỉ,Thẳng (Straight),8 (CB240),"2,470",8,19.76,7.81  
ST-06,Cầu thang ST1,Con kê thép mũ,Ghế kê (Chair),10 (CB240),580,12,6.96,4.29  
TỔNG,,,,,,,,138.40 kg  
Module này cho phép AI Agent chỉ cần nhận các kích thước cơ bản từ bản vẽ kiến trúc/kết cấu (chiều cao tầng, chiều dài vế, bề rộng bậc, chiều dày bản sàn) là có thể tự động suy ra toàn bộ thông số bê tông, ván khuôn và bảng thống kê cốt thép của cầu thang.  

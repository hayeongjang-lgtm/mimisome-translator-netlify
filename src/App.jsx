import React from 'react';
import * as XLSX from 'xlsx';

const RAW_DATA = `{"cat":{"보톡스/윤곽주사":{"en":"Botox / Face Contouring Injection","jp":"ボトックス・輪郭注射","cn":"肉毒素 / 轮廓针","tw":"肉毒素 / 輪廓針","th":"โบท็อกซ์ / ฉีดแฟต","vn":"Botox / Tiêm thon gọn line khuôn mặt"},"필러":{"en":"Filler","jp":"フィラー","cn":"填充剂","tw":"填充劑","th":"ฟิลเลอร์","vn":"Filler"},"리프팅":{"en":"Lifting","jp":"リフトアップ","cn":"提升","tw":"提升","th":"ยกกระชับ","vn":"Nâng cơ"},"스킨부스터":{"en":"Skin Booster","jp":"スキンブースター","cn":"皮肤增强剂","tw":"皮膚增強劑","th":"สกินบูสเตอร์","vn":"Skin Booster"},"줄기세포":{"en":"Stem Cells","jp":"幹細胞","cn":"干细胞","tw":"幹細胞","th":"สเต็มเซลล์","vn":"Tế bào gốc"},"실리프팅":{"en":"Thread Lifting","jp":"糸リフト","cn":"线雕","tw":"線雕","th":"ร้อยไหม","vn":"Căng chỉ"},"피부레이저(색소/혈관/기미)":{"en":"Skin Laser Treatment (Pigmentation / Vascular / Melasma)","jp":"皮膚レーザー（色素・血管・シミ）","cn":"皮肤激光 (色素 / 血管 / 黄褐斑)","tw":"皮膚雷射 (色素 / 血管 / 黃褐斑)","th":"เลเซอร์ผิว (เม็ดสี / รอยแดง / ฝ้า)","vn":"Laser da (sắc tố / mạch máu / nám)"},"여드름/모공/흉터":{"en":"Acne / Pores / Scars","jp":"ニキビ・毛穴・傷跡","cn":"痤疮 / 毛孔 / 疤痕","tw":"痤瘡 / 毛孔 / 疤痕","th":"สิว / รูขุมขน / หลุมสิว","vn":"Mụn trứng cá / Lỗ chân lông / Sẹo"},"제모":{"en":"Hair Removal","jp":"脱毛","cn":"脱毛","tw":"脫毛","th":"กำจัดขน","vn":"Triệt lông"},"비만/다이어트":{"en":"Weight Loss / Obesity Clinic","jp":"肥満・ダイエット","cn":"肥胖 / 减肥","tw":"肥胖 / 減肥","th":"ลดน้ำหนัก / รูปร่าง","vn":"Béo phì / Giảm cân"},"성형수술":{"en":"Cosmetic Surgery","jp":"美容整形手術","cn":"整形手术","tw":"整形手術","th":"ศัลยกรรม","vn":"Phẫu thuật thẩm mỹ"},"스킨케어":{"en":"Skin Care","jp":"スキンケア","cn":"皮肤护理","tw":"皮膚護理","th":"ดูแลผิว","vn":"Chăm sóc da"},"수액":{"en":"IV Drip","jp":"点滴","cn":"营养点滴","tw":"營養點滴","th":"วิตามินดริป","vn":"Truyền dịch dinh dưỡng"}},"tm":[{"k":"표정주름보톡스 1부위 (국산 하이톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) – Hitox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"jp":{"e":"表情じわボトックスの1部位（韓国製のハイトックス）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Hitox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Hitox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vn":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Hitox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (국산 코어톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Coretox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"jp":{"e":"表情じわボトックスの1部位（韓国製のコアトックス）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Coretox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Coretox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vn":{"e":"Botox nếp nhăn vùng mặt (Coretox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (국산 이노톡스)","kd":"세계 유일 액상 보톡스, 이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Innotox (Korean)","d":"World's only liquid Botox,  Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"jp":{"e":"表情じわボトックスの1部位（韓国製のイノトックス）","d":"世界唯一の液状型ボトックス おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Innotox）","d":"世界唯一液态肉毒素, 额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Innotox)","d":"โบท็อกซ์ชนิดน้ำหนึ่งเดียวในโลก, เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vn":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Innotox nội địa)","d":"Botox dạng lỏng duy nhất trên thế giới , Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (수입 제오민)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Xeomin (Imported)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"jp":{"e":"表情じわボトックスの1部位（海外製のゼオミン）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（进口 Xeomin）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (นำเข้า Xeomin)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vn":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Xeomin nhập khẩu)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"사각턱보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox – Hitox (Korean)"},"jp":{"e":"エラボトックス（韓国製のハイトックス）"},"cn":{"e":"方下巴肉毒素（韩国产 Hitox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Hitox)"},"vn":{"e":"Botox hàm vuông (Hitox nội địa)"}}},{"k":"사각턱보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox Coretox (Korean)"},"jp":{"e":"エラボトックス（韓国製のコアトックス）"},"cn":{"e":"方下巴肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Coretox)"},"vn":{"e":"Botox hàm vuông (Coretox nội địa)"}}},{"k":"사각턱보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox Innotox (Korean)"},"jp":{"e":"エラボトックス（韓国製のイノトックス）"},"cn":{"e":"方下巴肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Innotox)"},"vn":{"e":"Botox hàm vuông (Innotox nội địa)"}}},{"k":"사각턱보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox Xeomin (Imported)"},"jp":{"e":"エラボトックス（海外製のゼオミン）"},"cn":{"e":"方下巴肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์กราม (นำเข้า Xeomin)"},"vn":{"e":"Botox hàm vuông (Xeomin nhập khẩu)"}}},{"k":"풀페이스 스킨보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox – Hitox (Korean)"},"jp":{"e":"フルフェイススキンボトックス（韓国製のハイトックス）"},"cn":{"e":"全脸皮肤肉毒素（韩国产 Hitox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Hitox)"},"vn":{"e":"Skin Botox toàn mặt (Hitox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Coretox (Korean)"},"jp":{"e":"フルフェイススキンボトックス（韓国製のコアトックス）"},"cn":{"e":"全脸皮肤肉毒素（韩国产 Coretox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Coretox)"},"vn":{"e":"Skin Botox toàn mặt (Coretox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Innotox (Korean)"},"jp":{"e":"フルフェイススキンボトックス（韓国製のイノトックス）"},"cn":{"e":"全脸皮肤肉毒素（韩国产 Innotox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Innotox)"},"vn":{"e":"Skin Botox toàn mặt (Innotox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Xeomin (Imported)"},"jp":{"e":"フルフェイススキンボトックス（海外製のゼオミン）"},"cn":{"e":"全脸皮肤肉毒素（进口 Xeomin）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (นำเข้า Xeomin)"},"vn":{"e":"Skin Botox toàn mặt (Xeomin nhập khẩu)"}}},{"k":"승모근/종아리보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox – Hitox (Korean)"},"jp":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のハイトックス）"},"cn":{"e":"斜方肌/小腿肉毒素（韩国产 Hitox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Hitox)"},"vn":{"e":"Botox cơ thang/cơ bắp chân (Hitox nội địa)"}}},{"k":"승모근/종아리보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Coretox (Korean)"},"jp":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のコアトックス）"},"cn":{"e":"斜方肌/小腿肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Coretox)"},"vn":{"e":"Botox cơ thang/cơ bắp chân (Coretox nội địa)"}}},{"k":"승모근/종아리보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Innotox (Korean)"},"jp":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のイノトックス）"},"cn":{"e":"斜方肌/小腿肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Innotox)"},"vn":{"e":"Botox cơ thang/cơ bắp chân (Innotox nội địa)"}}},{"k":"승모근/종아리보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Xeomin (Imported)"},"jp":{"e":"僧帽筋・ふくらはぎボトックス（海外製のゼオミン）"},"cn":{"e":"斜方肌/小腿肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (นำเข้า Xeomin)"},"vn":{"e":"Botox cơ thang/cơ bắp chân (Xeomin nhập khẩu)"}}},{"k":"다한증보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox – Hitox (Korean)"},"jp":{"e":"多汗症ボトックス（韓国製のハイトックス）"},"cn":{"e":"多汗症肉毒素（韩国产 Hitox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Hitox)"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (Hitox nội địa)"}}},{"k":"다한증보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Coretox (Korean)"},"jp":{"e":"多汗症ボトックス（韓国製のコアトックス）"},"cn":{"e":"多汗症肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Coretox)"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (Coretox nội địa)"}}},{"k":"다한증보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Innotox (Korean)"},"jp":{"e":"多汗症ボトックス（韓国製のイノトックス）"},"cn":{"e":"多汗症肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Innotox)"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (Innotox nội địa)"}}},{"k":"다한증보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Xeomin (Imported)"},"jp":{"e":"多汗症ボトックス（海外製のゼオミン）"},"cn":{"e":"多汗症肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (นำเข้า Xeomin)"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (Xeomin nhập khẩu)"}}},{"k":"하관 대용량 양악주사","c":"보톡스/윤곽주사","t":{"en":{"e":"Lower Face Contouring Fat-Dissolving Injection"},"jp":{"e":"フェイスラインの大容量輪郭注射"},"cn":{"e":"下脸部大容量双颌针"},"th":{"e":"ฉีดหน้าเรียว แบบปริมาณมาก"},"vn":{"e":"Tiêm hàm dưới liều dung lượng cao"}}},{"k":"코조각주사","kd":"-","c":"보톡스/윤곽주사","t":{"en":{"e":"Nose Contouring Injection","d":"-"},"jp":{"e":"鼻彫刻注射","d":"-"},"cn":{"e":"鼻部雕刻针","d":"-"},"th":{"e":"ฉีดลดปีกจมูก","d":"-"},"vn":{"e":"Tiêm tạo dáng mũi"}}},{"k":"볼륨필러 (국산 아띠에르) 1cc","c":"필러","t":{"en":{"e":"Volume Filler 1cc – Attiére (Korean)"},"jp":{"e":"ボリュームアップフィラー（韓国製のアティエル）1㏄"},"cn":{"e":"丰盈填充（韩国产 Attiére）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (เกาหลี Attiére) 1cc"},"vn":{"e":"Filler tạo phồng (Attiére nội địa) 1cc"}}},{"k":"볼륨필러 (국산 순수필) 1cc","c":"필러","t":{"en":{"e":"Soonsoo Fill 1cc (Korean)"},"jp":{"e":"ボリュームアップフィラー（韓国製のスンスフィル）1㏄"},"cn":{"e":"丰盈填充（韩国产 Soonsoo Fill）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (เกาหลี Soonsoo Fill) 1cc"},"vn":{"e":"Filler tạo phồng (Soonsoo Fill nội địa) 1cc"}}},{"k":"볼륨필러 (수입 아말리안) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Amalian (Imported) 1cc"},"jp":{"e":"ボリュームアップフィラー（海外製のアマリアン）1㏄"},"cn":{"e":"丰盈填充（进口 Amalian）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (นำเข้า Amalian) 1cc"},"vn":{"e":"Filler tạo phồng (Amalian nhập khẩu) 1cc"}}},{"k":"볼륨필러 (수입 쥬비덤) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Juvederm (Imported) 1cc"},"jp":{"e":"ボリュームアップフィラー（海外製のジュビダーム）1㏄"},"cn":{"e":"丰盈填充（进口 Juvederm）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (นำเข้า Juvederm) 1cc"},"vn":{"e":"Filler tạo phồng (Juvederm nhập khẩu) 1cc"}}},{"k":"입술필러 무제한 (국산 아띠에르)","kd":"입꼬리 포함","c":"필러","t":{"en":{"e":"Unlimited Lip Filler – Attiére (Korean)","d":"Includes Mouth Corner Lift"},"jp":{"e":"唇フィラー無制限（韓国製のアティエル）","d":"口角を含む"},"cn":{"e":"嘴唇填充 不限容量（韩国产Attiére）","d":"包含嘴角"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc (เกาหลี Attiére)","d":"(รวมยกมุมปาก)"},"vn":{"e":"Filler môi không giới hạn (Attiére nội địa)","d":"Bao gồm cả khóe miệng"}}},{"k":"입술필러 무제한 (국산 순수필)","c":"필러","t":{"en":{"e":"Soonsoo Fill (Korean)"},"jp":{"e":"唇フィラー無制限（韓国製のスンスフィル）"},"cn":{"e":"嘴唇填充 不限容量（韩国产 Soonsoo Fill）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc(เกาหลี Soonsoo Fill)"},"vn":{"e":"Filler môi không giới hạn (Soonsoo Fill nội địa)"}}},{"k":"입술필러 무제한 (수입 아말리안)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Amalian (Imported)"},"jp":{"e":"唇フィラー無制限（海外製のアマリアン）"},"cn":{"e":"嘴唇填充 不限容量（进口 Amalian）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc(นำเข้า Amalian)"},"vn":{"e":"Filler môi không giới hạn (Amalian nhập khẩu) 1cc"}}},{"k":"입술필러 무제한 (수입 레스틸렌키스)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Restylane Kysse (Imported)"},"jp":{"e":"唇フィラー無制限（海外製のレスチレンキス）"},"cn":{"e":"嘴唇填充 不限容量（进口 Restylane Kysse）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc (นำเข้า Restylane Kysse)"},"vn":{"e":"Filler môi không giới hạn (Restylane Kysse nhập khẩu)"}}},{"k":"귀필러 (국산 아띠에르) 1cc","c":"필러","t":{"en":{"e":"Ear Filler 1cc – Attiére (Korean)"},"jp":{"e":"立ち耳フィラー（韓国製のアティエル）1㏄"},"cn":{"e":"耳部填充（韩国产 Attiére）1cc"},"th":{"e":"ฟิลเลอร์หู (เกาหลี Attiére) 1cc"},"vn":{"e":"Filler tai (Attiére nội địa) 1cc"}}},{"k":"귀필러 (수입 쥬비덤) 1cc","c":"필러","t":{"en":{"e":"Ear Filler Juvederm (Imported) 1cc"},"jp":{"e":"立ち耳フィラー（海外製のジュビダーム）1㏄"},"cn":{"e":"耳部填充（进口 Juvederm）1cc"},"th":{"e":"ฟิลเลอร์หู (นำเข้า Juvederm) 1cc"},"vn":{"e":"Filler tai (Juvederm nhập khẩu) 1cc"}}},{"k":"눈밑 전용필러 (수입 레스틸렌 아이라이트)","c":"필러","t":{"en":{"e":"Under-Eye Filler – Restylane Eyelight (Imported)"},"jp":{"e":"目の下の専用フィラー（海外製のレスチレンアイライト）"},"cn":{"e":"眼下专用填充（进口 Restylane Eyelight）"},"th":{"e":"ฟิลเลอร์สำหรับใต้ตา (นำเข้า Restylane Eyelight)"},"vn":{"e":"Filler chuyên dụng vùng dưới mắt (nhập khẩu Restylane Eyelight)"}}},{"k":"바디필러 1cc (국산)","c":"필러","t":{"en":{"e":"Body Filler 1cc (Korean)"},"jp":{"e":"ボディフィラー1㏄（韓国製）"},"cn":{"e":"身体填充 1cc（韩国产）"},"th":{"e":"บอดี้ฟิลเลอร์ 1cc (เกาหลี)"},"vn":{"e":"Filler body 1cc (nội địa)"}}},{"k":"리니어지 100샷","c":"리프팅","t":{"en":{"e":"LinearZ 100 Shots"},"jp":{"e":"リニアージ100ショット"},"cn":{"e":"LinearZ 100发"},"th":{"e":"ลิเนียร์จี 100 ช็อต"},"vn":{"e":"LinearZ 100 shot"}}},{"k":"인모드fx 풀페이스","c":"리프팅","t":{"en":{"e":"InMode FX Full Face"},"jp":{"e":"インモードFXフルフェイス"},"cn":{"e":"InMode FX 全脸"},"th":{"e":"InMode FX ทั่วหน้า"},"vn":{"e":"InMode FX toàn mặt"}}},{"k":"인모드forma 풀페이스","c":"리프팅","t":{"en":{"e":"InMode FORMA Full Face"},"jp":{"e":"インモードFORMAフルフェイス"},"cn":{"e":"InMode FORMA 全脸"},"th":{"e":"InMode FORMA ทั่วหน้า"},"vn":{"e":"InMode FORMA  toàn mặt"}}},{"k":"덴서티 100샷","c":"리프팅","t":{"en":{"e":"Density 100 Shots"},"jp":{"e":"デンシティ100ショット"},"cn":{"e":"Density 100发"},"th":{"e":"Density 100 ช็อต"},"vn":{"e":"Density 100 shot"}}},{"k":"포텐자 DDR 500샷","c":"리프팅","t":{"en":{"e":"Potenza DDR 500 Shots"},"jp":{"e":"ポテンツァDDR500ショット"},"cn":{"e":"Potenza DDR 500发"},"th":{"e":"โพเทนซ่า DDR 500 ช็อต"},"vn":{"e":"Potenza DDR 500 shot"}}},{"k":"포텐자 DIA 500샷","c":"리프팅","t":{"en":{"e":"Potenza DIA 500 Shots"},"jp":{"e":"ポテンツァDIA500ショット"},"cn":{"e":"Potenza DIA 500发"},"th":{"e":"โพเทนซ่า DIA 500 ช็อต"},"vn":{"e":"Potenza DIA 500 shot"}}},{"k":"울쎄라피프라임 100샷","c":"리프팅","t":{"en":{"e":"Ultherapy Prime 100 Shots"},"jp":{"e":"ウルセラプライム100ショット"},"cn":{"e":"Ultherapy Prime 100发"},"th":{"e":"อัลเทอราไพรม์ 100 ช็อต"},"vn":{"e":"Ultherapy Prime 100 shot"}}},{"k":"써마지FLX 300샷","c":"리프팅","t":{"en":{"e":"Thermage FLX 300 Shots"},"jp":{"e":"サマージFLX300ショット"},"cn":{"e":"Thermage FLX 300发"},"th":{"e":"เทอร์มาจ FLX 300 ช็อต"},"vn":{"e":"Thermage FLX 300 shot"},"tw":{"e":"第四代鳳凰電波 300發"}}},{"k":"물광주사 하이주 2cc","c":"스킨부스터","t":{"en":{"e":"Hy lnj Skin Booster Injection 2cc"},"jp":{"e":"水光注射ハイジュ2㏄"},"cn":{"e":"水光针 High lnj 2cc"},"th":{"e":"Hy lnj สกินบูสเตอร์ 2cc"},"vn":{"e":"Tiêm căng bóng da Hy Inj 2cc"}}},{"k":"릴리이드M 5cc","c":"스킨부스터","t":{"en":{"e":"Lilied M 5cc"},"jp":{"e":"リリイドM5㏄"},"cn":{"e":"Lilied M 5cc"},"th":{"e":"Lilied M 5cc"},"vn":{"e":"Lilied M 5cc"}}},{"k":"리즈네 2cc","c":"스킨부스터","t":{"en":{"e":"Lizne 2cc"},"jp":{"e":"リズネ2㏄"},"cn":{"e":"Lizne 2cc"},"th":{"e":"Lizne 2cc"},"vn":{"e":"Lizne 2cc"}}},{"k":"아이 리쥬란 1cc","kd":"눈밑 전용","c":"스킨부스터","t":{"en":{"e":"Eye Rejuran 1cc","d":"For Under-Eye Area Only"},"jp":{"e":"リジュランアイ1㏄","d":"目の下専用"},"cn":{"e":"眼部 Rejuran 1cc","d":"眼底专用"},"th":{"e":"Eye Rejuran 1cc","d":"(สำหรับใต้ตาโดยเฉพาะ)"},"vn":{"e":"Rejuran Eye 1cc","d":"Chuyên dùng cho vùng mắt"}}},{"k":"리쥬란HB 1cc","c":"스킨부스터","t":{"en":{"e":"Rejuran HB 1cc"},"jp":{"e":"リジュランHB 1㏄"},"cn":{"e":"Rejuran HB 1cc"},"th":{"e":"Rejuran HB 1cc"},"vn":{"e":"Rejuran HB 1cc"}}},{"k":"리쥬란S 1cc","c":"스킨부스터","t":{"en":{"e":"Rejuran S 1cc"},"jp":{"e":"リジュランS1㏄"},"cn":{"e":"Rejuran S 1cc"},"th":{"e":"Rejuran S 1cc"},"vn":{"e":"Rejuran S 1cc"}}},{"k":"엑소좀 3D 네오스템 1병","kd":"항염효과 인증","c":"스킨부스터","t":{"en":{"e":"Exosome 3D NeoStem 1 Vial","d":"Certified Anti-Inflammatory Effect"},"jp":{"e":"ネオステム３Dエクソソーム1バイアル","d":"抗炎症作用認証"},"cn":{"e":"外泌体 3D NeoStem 1瓶","d":"抗炎效果认证"},"th":{"e":"Exosome 3D Neostem 1 ขวด","d":"(ผ่านการรับรองเรื่องลดการอักเสบ)"},"vn":{"e":"Exosome 3D Neostem 1 lọ","d":"Chứng nhận hiệu quả kháng viêm"}}},{"k":"쥬베룩 스킨부스터 4cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Skin Booster 4cc"},"jp":{"e":"ジュベルックスキンブースター4㏄"},"cn":{"e":"Juvelook 水光针 4cc"},"th":{"e":"Juvelook Skin Booster 4cc"},"vn":{"e":"Juvelook Skin Booster 4cc"},"tw":{"e":"Juvelook 水光針 4cc"}}},{"k":"쥬베룩 아이 1병","kd":"눈밑 전용","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial","d":"For Under-Eye Area Only"},"jp":{"e":"ジュベルックアイ1バイアル","d":"目の下専用"},"cn":{"e":"Juvelook 眼部 1瓶","d":"眼底专用"},"th":{"e":"Juvelook Eye 1 ขวด","d":"(สำหรับใต้ตาโดยเฉพาะ)"},"vn":{"e":"Juvelook Eye 1 lọ","d":"Chuyên dùng cho vùng mắt"},"tw":{"e":"Juvelook 眼部 1瓶"}}},{"k":"쥬베룩 볼륨 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Volume 1 Vial"},"jp":{"e":"ジュベルックボリューム1バイアル"},"cn":{"e":"Juvelook 丰盈 1瓶"},"th":{"e":"Juvelook Volume 1 ขวด"},"vn":{"e":"Juvelook Volume 1 lọ"},"tw":{"e":"Juvelook 豐盈 1瓶"}}},{"k":"쥬비덤 스킨바이브 1cc","c":"스킨부스터","t":{"en":{"e":"Juvederm Skinvive 1cc"},"jp":{"e":"スキンバイブbyジュビダーム1㏄"},"cn":{"e":"Juvederm Skinvive 1cc"},"th":{"e":"Juvederm Skinvive 1cc"},"vn":{"e":"Juvederm Skinvive 1cc"}}},{"k":"벨로테로 리바이브 1cc","c":"스킨부스터","t":{"en":{"e":"Belotero Revive 1cc"},"jp":{"e":"ベロテロリバイブ1㏄"},"cn":{"e":"Belotero Revive 1cc"},"th":{"e":"Belotero Revive 1cc"},"vn":{"e":"Belotero Revive 1cc"}}},{"k":"리투오 5cc","c":"스킨부스터","t":{"en":{"e":"Re2O 5cc"},"jp":{"e":"リトゥオ5㏄"},"cn":{"e":"Re2O 5cc"},"th":{"e":"Re2O 5cc"},"vn":{"e":"Re2O 5cc"}}},{"k":"셀르디엠 1시린지","c":"스킨부스터","t":{"en":{"e":"CellreDM 1 Syringe"},"jp":{"e":"セルディエム1シリンジ"},"cn":{"e":"CellreDM 1支针剂"},"th":{"e":"CellreDM 1 Syringe"},"vn":{"e":"CellreDM 1 ống tiêm"}}},{"k":"스컬트라 1병","c":"스킨부스터","t":{"en":{"e":"Sculptra – 1 Vial"},"jp":{"e":"スカルトラ1バイアル"},"cn":{"e":"Sculptra 1瓶"},"th":{"e":"Sculptra 1 ขวด"},"vn":{"e":"Sculptra 1 lọ"}}},{"k":"레디어스 1시린지","c":"스킨부스터","t":{"en":{"e":"Radiesse – 1 Syringe"},"jp":{"e":"レディエッセ1シリンジ"},"cn":{"e":"Radiesse 1支针剂"},"th":{"e":"Radiesse 1 Syringe"},"vn":{"e":"Radiesse 1 mũi tiêm"}}},{"k":"미라셀 줄기세포","c":"줄기세포","t":{"en":{"e":"Miracell Stem Cell Therapy"},"jp":{"e":"ミラセル幹細胞"},"cn":{"e":"Miracell 干细胞"},"th":{"e":"Miracell Stem Cell"},"vn":{"e":"Miracell tế bào gốc"}}},{"k":"울트라브이 콜라겐실 1줄","c":"실리프팅","t":{"en":{"e":"Ultra V Collagen Thread – 1 Thread"},"jp":{"e":"ウルトラブイのコラーゲン糸1本"},"cn":{"e":"Ultra V 胶原蛋白线 1根"},"th":{"e":"ไหมคอลลาเจน Ultra V 1 เส้น"},"vn":{"e":"Sợi chỉ Collagen Ultra V 1 sợi"}}},{"k":"실리프팅 1줄","kd":"심부볼,팔자,턱선 리프팅","c":"실리프팅","t":{"en":{"e":"Thread Lifting – 1 Thread","d":"Deep Cheek, Nasolabial Fold & Jawline Lifting"},"jp":{"e":"糸リフト1本","d":"バッカルファット、ほうれい線、 顎のラインリフトアップ"},"cn":{"e":"线雕 1根","d":"深层脸颊脂肪、法令纹、下颌线提升"},"th":{"e":"ร้อยไหม 1 เส้น","d":"(ยกกระพุ้งแก้ม ร่องแก้ม กรอบหน้า)"},"vn":{"e":"Căng chỉ 1 sợi","d":"Nâng cơ má hóp, rãnh cười, đường viền hàm"}}},{"k":"잼버실 2줄","kd":"팔자,볼,인디언주름 볼륨","c":"실리프팅","t":{"en":{"e":"Jamber Thread Lift – 2 Threads","d":"Nasolabial Fold / Cheek / Indian Fold Volume Enhancement"},"jp":{"e":"ジャンバー糸2本","d":"ほうれい線、頬、インディアンラインボリューム"},"cn":{"e":"Jamber线2根","d":"法令纹、脸颊、印第安纹容量填充"},"th":{"e":"ไหม Jamber 2 เส้น","d":"(เพิ่มวอลุ่มร่องแก้ม กระพุ้งแก้ม ร่องอินเดียน)"},"vn":{"e":"Chỉ Jamber 2 sợi","d":"Tăng thể tích rãnh cười, má và rãnh lệ có nếp nhăn"}}},{"k":"이마거상,폭스아이 실리프팅 1줄","c":"실리프팅","t":{"en":{"e":"Brow Lift / Fox Eye Thread Lift – 1 Thread"},"jp":{"e":"おでこリフト、フォックスアイ糸リフト1本"},"cn":{"e":"额头提升、Fox Eye线雕 1根"},"th":{"e":"ร้อยไหมยกหน้าผาก / Fox Eye 1 เส้น"},"vn":{"e":"Căng chỉ trán, mắt cáo 1 sợi"}}},{"k":"콧대주름 실리프팅 1줄","c":"실리프팅","t":{"en":{"e":"Nose Bridge Wrinkle Thread Lift – 1 Thread"},"jp":{"e":"バニーラインの糸リフト1本"},"cn":{"e":"鼻梁皱纹线雕 1根"},"th":{"e":"ร้อยไหมสันจมูก 1 เส้น"},"vn":{"e":"Nâng chỉ vùng nếp nhăn sống mũi 1 sợi"}}},{"k":"이중턱 근육묶기 실리프팅 2줄","c":"실리프팅","t":{"en":{"e":"Double Chin Muscle-Tightening Thread Lift – 2 Threads"},"jp":{"e":"二十顎の筋肉縛りの糸リフト2本"},"cn":{"e":"双下巴肌肉固定线雕 2根"},"th":{"e":"ร้อยไหมเก็บเหนียง 2 เส้น"},"vn":{"e":"Căng chỉ nâng cơ thu gọn cằm hai ngấn 2 sợi"}}},{"k":"점,비립종,사마귀제거 1개","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Mole / Milia / Wart Removal – 1 Lesion"},"jp":{"e":"ほくろ、稗粒腫、いぼの除去1個"},"cn":{"e":"痣、脂肪粒、去跖疣 1个"},"th":{"e":"กำจัดไฝ / สิวข้าวสาร / หูด 1 จุด"},"vn":{"e":"Loại bỏ nốt ruồi, mụn kê, hoặc mụn cóc 1 nốt"}}},{"k":"듀얼토닝(브라이톤+제네시스) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Dual Toning (Brightone + Genesis) – 1 Session"},"jp":{"e":"デュアルトーニング（ブライトーン＋ジェネシス）1回"},"cn":{"e":"双重嫩肤（Brightone + Genesis）1次"},"th":{"e":"ดูอัลโทนนิ่ง (Brightone + Genesis) 1 ครั้ง"},"vn":{"e":"Dual Toning (Brightone + Genesis) 1 lần"}}},{"k":"피코토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Pico Toning – 1 Session"},"jp":{"e":"ピコトーニング1回"},"cn":{"e":"Pico嫩肤 1次"},"th":{"e":"พิโคโทนนิ่ง 1 ครั้ง"},"vn":{"e":"Pico Toning 1 lần"}}},{"k":"헐리우드 스펙트라토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Toning – 1 Session"},"jp":{"e":"ハリウッドスペクトラ1回"},"cn":{"e":"Hollywood Spectra嫩肤 1次"},"th":{"e":"Hollywood Spectra Toning 1 ครั้ง"},"vn":{"e":"Hollywood Spectra Toning 1 lần"}}},{"k":"엑셀브이 플러스(홍조/색소) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Excel V Plus (Redness/Pigmentation) – 1 Session"},"jp":{"e":"エクセルVプラス（紅潮・色素）１回"},"cn":{"e":"Excel V Plus（红血丝/色素）1次"},"th":{"e":"Excel V พลัส (รอยแดง / เม็ดสี) 1 ครั้ง"},"vn":{"e":"Excel V Plus (điều trị mẩn đỏ/ sắc tố) 1 lần"}}},{"k":"풀페이스 혈관레이저 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Full Face Vascular Laser – 1 Session"},"jp":{"e":"フルフェイス血管レーザー１回"},"cn":{"e":"全脸血管激光 1次"},"th":{"e":"เลเซอร์เส้นเลือดทั่วหน้า 1 ครั้ง"},"vn":{"e":"Laser mạch máu toàn mặt 1 lần"}}},{"k":"포텐자(기미) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Potenza (Melasma) – 1 Session"},"jp":{"e":"ポテンツァ（シミ）1回"},"cn":{"e":"Potenza（黄褐斑）1次"},"th":{"e":"โพเทนซ่า ฝ้า 1 ครั้ง"},"vn":{"e":"Potenza (điều trị nám) 1 lần"}}},{"k":"얼굴전체 잡티레이저(클라리티) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Full Face Pigmentation Laser (Clarity) – 1 Session"},"jp":{"e":"顔全体のシミ・くすみレーザー（クラリティ）1回"},"cn":{"e":"全脸色斑激光（Clarity）1次"},"th":{"e":"เลเซอร์กำจัดจุดด่างดำทั่วหน้า (Clarity) 1 ครั้ง"},"vn":{"e":"Laser đốm sắc tố toàn mặt (Clarity) – 1 lần"}}},{"k":"리팟 흑자제거 1부위","kd":"1cm 이내","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"REEPOT Age Spot Removal – 1 Area","d":"Within 1cm"},"jp":{"e":"リポットレーザー黒子除去1部位","d":"1cm 이내　1㎝以内"},"cn":{"e":"REEPOT 黑斑去除 1部位","d":"1cm以内"},"th":{"e":"REEPOT กำจัดกระแดด 1 บริเวณ","d":"(ไม่เกิน 1 ซม.)"},"vn":{"e":"REEPOT loại bỏ đốm nâu 1 vùng","d":"Trong phạm vi 1cm"}}},{"k":"눈썹 문신제거 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Eyebrow Tattoo Removal – 1 Session"},"jp":{"e":"眉毛のアートメイク除去1回"},"cn":{"e":"眉毛纹身去除 1次"},"th":{"e":"ลบรอยสักคิ้ว 1 ครั้ง"},"vn":{"e":"Xóa xăm chân mày 1 lần"}}},{"k":"아이라인 문신제거 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Eyeliner Tattoo Removal – 1 Session"},"jp":{"e":"アイラインのアートメイク除去1回"},"cn":{"e":"眼线纹身去除 1次"},"th":{"e":"ลบรอยสักอายไลเนอร์ 1 ครั้ง"},"vn":{"e":"Xóa xăm mí mắt 1 lần"}}},{"k":"흑백,컬러,레터링 문신제거","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Black & Gray / Color / Lettering Tattoo Removal"},"jp":{"e":"黒と白、カラー、レタリングタトゥー除去"},"cn":{"e":"黑白、彩色、文字纹身去除"},"th":{"e":"ลบรอยสักขาวดำ / สี / ตัวอักษร"},"vn":{"e":"Xóa xăm đơn sắc, màu hoặc Lettering"}}},{"k":"애플토닝 1회","kd":"피지선 억제","c":"여드름/모공/흉터","t":{"en":{"e":"Apple Toning – 1 Session","d":"Sebaceous Gland Suppression"},"jp":{"e":"アップルトーニング1回","d":"皮脂腺の抑制"},"cn":{"e":"Apple嫩肤 1次","d":"抑制皮脂腺"},"th":{"e":"แอปเปิลโทนนิ่ง 1 ครั้ง","d":"(ช่วยลดการทำงานของต่อมไขมัน)"},"vn":{"e":"Apple Toning 1 lần","d":"Ức chế tuyến bã nhờn"}}},{"k":"포텐자(모공/여드름) 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza (Pores/Acne) – 1 Session"},"jp":{"e":"ポテンツァ（毛穴・ニキビ）1回"},"cn":{"e":"Potenza（毛孔/痘痘）1次"},"th":{"e":"โพเทนซ่า รูขุมขน / สิว 1 ครั้ง"},"vn":{"e":"Potenza (lỗ chân lông / mụn trứng cá) 1 lần"}}},{"k":"어븀프락셀 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Erbium Fraxel – 1 Session"},"jp":{"e":"エルビウムフラクショナルレーザー1回"},"cn":{"e":"Erbium Fraxel 1次"},"th":{"e":"เออร์เบียมแฟรกเซล 1 ครั้ง"},"vn":{"e":"Erbium Fraxel 1 lần"}}},{"k":"피코프락셀 MLA 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel MLA – 1 Session"},"jp":{"e":"ピコフラクショナルレーザーMLA1回"},"cn":{"e":"Pico Fraxel MLA 1次"},"th":{"e":"พิโคแฟรกเซล MLA 1 ครั้ง"},"vn":{"e":"Pico Fraxel MLA 1 lần"}}},{"k":"남성 인중제모 1회","c":"제모","t":{"en":{"e":"Men’s Upper Lip Hair Removal – 1 Session"},"jp":{"e":"男性鼻下脱毛1回"},"cn":{"e":"男性 人中脱毛 1次"},"th":{"e":"กำจัดขนหนวด (ชาย) 1 ครั้ง"},"vn":{"e":"Triệt ria mép nam 1 lần"}}},{"k":"남성 인중,앞턱제모 1회","c":"제모","t":{"en":{"e":"Men’s Upper Lip & Chin Hair Removal – 1 Session"},"jp":{"e":"男性鼻下、顎脱毛1回"},"cn":{"e":"男性 人中、下巴前部脱毛 1次"},"th":{"e":"กำจัดขนหนวดและเครา (ชาย) 1 ครั้ง"},"vn":{"e":"Triệt ria mép và cằm trước nam 1 lần"}}},{"k":"남성 산타제모 1회","kd":"인중,턱수염,구렛나루,볼","c":"제모","t":{"en":{"e":"Men’s “Santa” Beard Hair Removal – 1 Session","d":"Upper Lip / Beard / Sideburns / Cheeks"},"jp":{"e":"男性サンタ脱毛1回","d":"鼻下、顎、もみあげ、頬"},"cn":{"e":"男性 Santa脱毛 1次","d":"人中、胡须、鬓角、脸颊"},"th":{"e":"กำจัดขนบริเวณหนวดและแก้ม 1 ครั้ง","d":"หนวด, เครา, จอน, แก้ม"},"vn":{"e":"Triệt râu quai nón nam 1 lần","d":"Ria mép, râu cằm, tóc mai, má"}}},{"k":"남성 수염전체제모 1회","kd":"인중,턱수염,구렛나루,볼,목라인","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal – 1 Session","d":"Includes Neck Line"},"jp":{"e":"男性ヒゲ全部位脱毛1回","d":"鼻下、顎、もみあげ、頬、首"},"cn":{"e":"男性 全胡须脱毛 1次","d":"人中、胡须、鬓角、脸颊、颈部线条"},"th":{"e":"กำจัดขนทั่วใบหน้าและลำคอ (ชาย) 1 ครั้ง","d":"หนวด, เครา, จอน, แก้ม, แนวลำคอ"},"vn":{"e":"Triệt toàn bộ râu nam 1 lần","d":"Vùng cổ"}}},{"k":"여성 인중제모 1회","c":"제모","t":{"en":{"e":"Women’s Upper Lip Hair Removal – 1 Session"},"jp":{"e":"女性鼻下脱毛1回"},"cn":{"e":"女性 人中脱毛 1次"},"th":{"e":"กำจัดขนหนวด (หญิง) 1 ครั้ง"},"vn":{"e":"Triệt ria mép nữ 1 lần"}}},{"k":"여성 겨드랑이제모 1회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal – 1 Session"},"jp":{"e":"女性両ワキ脱毛1回"},"cn":{"e":"女性 腋下脱毛 1次"},"th":{"e":"กำจัดขนรักแร้ (หญิง) 1 ครั้ง"},"vn":{"e":"Triệt lông nách nữ 1 lần"}}},{"k":"여성 종아리제모 1회","kd":"무릎포함","c":"제모","t":{"en":{"e":"Women’s Lower Leg Hair Removal – 1 Session","d":"Includes Knees"},"jp":{"e":"女性ふくらはぎ脱毛1回","d":"膝を含む"},"cn":{"e":"女性 小腿脱毛 1次","d":"包含膝盖"},"th":{"e":"กำจัดขนหน้าแข้ง (หญิง) 1 ครั้ง","d":"รวมเข่า"},"vn":{"e":"Triệt lông bắp chân nữ 1 lần","d":"Bao gồm cả đầu gối"}}},{"k":"바디 리니어지 100샷","c":"비만/다이어트","t":{"en":{"e":"Body LinearZ 100 Shots"},"jp":{"e":"ボディリニアージ100ショット"},"cn":{"e":"Body LinearZ 100发"},"th":{"e":"Body LinearZ 100 ช็อต"},"vn":{"e":"Body LinearZ 100 shot"}}},{"k":"지방분해주사 용량무제한 1회","c":"비만/다이어트","t":{"en":{"e":"Unlimited Fat-Dissolving Injection – 1 Session"},"jp":{"e":"脂肪溶解注射・容量無制限1回"},"cn":{"e":"溶脂针 不限容量 1次"},"th":{"e":"ฉีดสลายไขมัน (ไม่จำกัดปริมาณ) 1 ครั้ง"},"vn":{"e":"Tiêm tan mỡ, không giới hạn liều, 1 lần"}}},{"k":"지방추출주사 에스펫 1병","kd":"50cc","c":"비만/다이어트","t":{"en":{"e":"S-FAT Fat Extraction Injection – 1 Vial","d":"50cc"},"jp":{"e":"脂肪吸引注射エスファット1瓶","d":"50cc"},"cn":{"e":"抽脂针 S- FAT 1瓶","d":"50cc"},"th":{"e":"S-FAT ดูดไขมัน (1 ขวด)","d":"50cc"},"vn":{"e":"Tiêm chiết xuất mỡ S-FAT 1 lọ","d":"50cc"}}},{"k":"위고비","c":"비만/다이어트","t":{"en":{"e":"Wegovy"},"jp":{"e":"ウゴービ"},"cn":{"e":"Wegovy"},"th":{"e":"Wegovy"},"vn":{"e":"Wegovy"}}},{"k":"보조개 한쪽","c":"성형수술","t":{"en":{"e":"Dimple Creation – One Side"},"jp":{"e":"えくぼ形成片側"},"cn":{"e":"酒窝（单侧）"},"th":{"e":"ทำลักยิ้ม (1 ข้าง)"},"vn":{"e":"Một bên lúm đồng tiền"}}},{"k":"풀페이스 지방이식","kd":"수면마취 가능","c":"성형수술","t":{"en":{"e":"Full Face Fat Grafting","d":"Sedation Anesthesia Available"},"jp":{"e":"フルフェイス脂肪注入","d":"睡眠麻酔可能"},"cn":{"e":"全脸脂肪移植","d":"可进行睡眠麻醉"},"th":{"e":"เติมไขมันทั่วใบหน้า (Fat Grafting)","d":"หลับด้วยยาสลบได้"},"vn":{"e":"Cấy mỡ toàn mặt","d":"Có thể gây mê ngủ"}}},{"k":"눈밑 지방제거+재배치+눈밑,앞볼 지방이식","kd":"수면마취 가능","c":"성형수술","t":{"en":{"e":"Under-Eye Fat Removal + Fat Repositioning + Under-Eye & Front Cheek Fat Grafting"},"jp":{"e":"目の下の脂肪取り＋再配置＋目の下、頬の脂肪注入"},"cn":{"e":"眼下脂肪去除+重新排列+眼下、前脸颊脂肪移植"},"th":{"e":"จัดเรียงไขมันใต้ตา + เติมไขมันใต้ตาและแก้ม"},"vn":{"e":"Lấy mỡ bọng mắt + tái phân bố + cấy mỡ vùng dưới mắt và má trước"}}},{"k":"미백/보습/진정/재생관리 1회","c":"성형수술","t":{"en":{"e":"Brightening / Hydration / Soothing / Regeneration Care – 1 Session"},"jp":{"e":"美白・保湿・鎮静・再生管理1回"},"cn":{"e":"美白/保湿/舒缓/再生护理 1次"},"th":{"e":"โปรแกรมกระจ่างใส/เติมความชุ่มชื้น/ปลอบประโลมผิว 1 ครั้ง"},"vn":{"e":"Chăm sóc làm trắng/giữ ẩm/làm dịu/tái tạo 1 lần"}}},{"k":"아쿠아필+진정관리 1회","c":"성형수술","t":{"en":{"e":"Aqua Peel + Soothing Care – 1 Session"},"jp":{"e":"アクアピール＋鎮静管理1回"},"cn":{"e":"Aqua Peel + 舒缓护理 1次"},"th":{"e":"อาควาพีล + ทรีทเม้นท์ปลอบประโลมผิว 1 ครั้ง"},"vn":{"e":"Aqua Peel + chăm sóc làm dịu 1 lần"}}},{"k":"LDM 물방울 초음파관리 1회","c":"성형수술","t":{"en":{"e":"LDM Water-Drop Ultrasound Treatment – 1 Session"},"jp":{"e":"LDM水玉超音波管理1回"},"cn":{"e":"LDM 水滴超声波护理 1次"},"th":{"e":"LDM Water Drop Lifting 1 ครั้ง"},"vn":{"e":"Chăm sóc da bằng sóng siêu âm LDM Water-Drop 1 lần"}}},{"k":"오투덤 산소테라피 1회","c":"성형수술","t":{"en":{"e":"O2Derm Oxygen Therapy – 1 Session"},"jp":{"e":"オーツーダム酸素テラピー1回"},"cn":{"e":"O2Derm 氧气疗法 1次"},"th":{"e":"O2Derm Oxygen Therapy 1 ครั้ง"},"vn":{"e":"O2Derm Oxygen Therapy 1 lần"}}},{"k":"4세대 라라필 1회","c":"성형수술","t":{"en":{"e":"4th Generation LHALA Peel – 1 Session"},"jp":{"e":"第4世代のララピール1回"},"cn":{"e":"第四代 LHALA Peel 1次"},"th":{"e":"LHALA Peel (Gen 4) 1 ครั้ง"},"vn":{"e":"LHALA Peel thế hệ 4 – 1 lần"}}},{"k":"백옥/신데렐라주사 1회","c":"성형수술","t":{"en":{"e":"Glutathione / Cinderella IV Drip – 1 Session"},"jp":{"e":"白玉・シンデレラ注射1回"},"cn":{"e":"白玉/灰姑娘针 1次"},"th":{"e":"ฉีดกลูต้า/ฉีดซินเดอเรลล่า 1 ครั้ง"},"vn":{"e":"Tiêm Bạch Ngọc hoặc Cinderella 1 lần"}}},{"k":"항노화주사 1회","kd":"피로 회복,항산화","c":"성형수술","t":{"en":{"e":"Anti-Aging IV Drip – 1 Session","d":"Fatigue Recovery & Antioxidant Effect"},"jp":{"e":"抗老化注射1回","d":"疲労回復、抗酸化"},"cn":{"e":"抗衰老针 1次","d":"疲劳恢复、抗氧化"},"th":{"e":"ฉีดชะลอวัย 1 ครั้ง","d":"ฟื้นฟูร่างกาย, ต้านอนุมูลอิสระ"},"vn":{"e":"Tiêm chống lão hóa 1 lần","d":"Đỡ mệt mỏi, chống oxy hóa"}}},{"k":"상쾌한주사 1회","kd":"숙취 해소,피로 회복","c":"성형수술","t":{"en":{"e":"Energy Recovery IV Drip – 1 Session","d":"Hangover Relief & Fatigue Recovery"},"jp":{"e":"爽快注射1回","d":"二日酔い解消、疲労回復"},"cn":{"e":"清爽针 1次","d":"解酒、疲劳恢复"},"th":{"e":"ฉีดแก้แฮงค์ 1 ครั้ง","d":"แก้แฮงค์, ฟื้นฟูร่างกาย"},"vn":{"e":"Tiêm phục hồi sảng khoái 1 lần","d":"Giải rượu, hồi phục mệt mỏi"}}},{"k":"신비주사(신데렐라+비타민) 1회","kd":"노화 예방,체지방 감소,활성산소 제거,면역강화","c":"성형수술","t":{"en":{"e":"Shinbi IV Drip (Cinderella + Vitamins) – 1 Session","d":"Anti-Aging / Body Fat Reduction / Antioxidant Effect / Immune Support"},"jp":{"e":"シンビ注射（シンデレラ＋ビタミン）1回","d":"老化予防、体脂肪減少、活性酸素除去、免疫強化"},"cn":{"e":"神秘针（灰姑娘+维他命）1次","d":"抗衰老、减少体脂、清除活性氧、增强免疫力"},"th":{"e":"ฉีด Shinbi (ฉีดซินเดอเรลล่า+ฉีดวิตามิน) 1 ครั้ง","d":"คุมน้ำหนัก, เสริมภูมิคุ้มกัน"},"vn":{"e":"Tiêm Shinbi (Cinderella + Vitamin) 1 lần","d":"Ngăn ngừa lão hóa, giảm mỡ cơ thể, loại bỏ oxy hoạt tính, tăng cường miễn dịch"}}},{"k":"오로라주사(백옥+비타민) 1회","kd":"노화 예방,피부톤 개선,간 해독,활성산소 제거","c":"성형수술","t":{"en":{"e":"Aurora IV Drip (Glutathione + Vitamins) – 1 Session","d":"Anti-Aging / Skin Tone Improvement / Liver Detox / Antioxidant Effect"},"jp":{"e":"オーロラ注射（白玉＋ビタミン）1回","d":"老化予防、肌のトーンアップ、肝臓解毒、活性酸素除去"},"cn":{"e":"极光针（白玉+维他命）1次","d":"抗衰老、改善肤色、肝脏解毒、清除活性氧"},"th":{"e":"ฉีด Aurora(ฉีดกลูต้า+ฉีดวิตามิน) 1 ครั้ง","d":"ดีท็อกซ์ตับ, ปรับสีผิว"},"vn":{"e":"Tiêm Aurora (Bạch Ngọc + Vitamin) 1 lần","d":"Ngăn ngừa lão hóa, cải thiện tone da, thải độc gan, loại bỏ oxy hoạt tính"}}},{"k":"소녀주사(태반+백옥) 1회","kd":"갱년기증상 완화,피부톤 개선,저속노화","c":"성형수술","t":{"en":{"e":"Girl IV Drip (Placenta + Glutathione) – 1 Session","d":"Menopause Symptom Relief / Skin Tone Improvement / Slow Aging Support"},"jp":{"e":"少女注射（プラセンタ＋白玉）1回","d":"更年期症状改善、肌のトーンアップ、スローエイジング"},"cn":{"e":"少女针（胎盘素+白玉）1次","d":"缓解更年期症状、改善肤色、延缓衰老"},"th":{"e":"ฉีด Girl(ฉีดรกเด็ก+ฉีดกลูต้า) 1 ครั้ง","d":"ปรับสมดุลวัยทอง, ชะลอวัย"},"vn":{"e":"Tiêm trẻ hóa (Nhau thai + Bạch Ngọc) 1 lần","d":"Giảm triệu chứng mãn kinh, cải thiện tone da, làm chậm lão hóa"}}},{"k":"표정주름보톡스 1부위 (국산 이노톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Innotox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Innotox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vn":{"e":"Botox nếp nhăn vùng mặt (Innotox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"},"jp":{"e":"表情じわボトックスの1部位（韓国製のイノトックス）","d":"世界唯一の液状型ボトックス\\nおでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Innotox)","d":"โบท็อกซ์ชนิดน้ำหนึ่งเดียวในโลก, เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"}}},{"k":"표정주름보톡스 1부위 (국산 제테마)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) JETEMA (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 JETEMA）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vn":{"e":"Botox nếp nhăn vùng mặt (JETEMA nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (수입 디스포트)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Dysport (Imported)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"cn":{"e":"表情肌肉皱纹肉毒素 1部位（进口 Dysport）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vn":{"e":"Botox nếp nhăn vùng mặt (Dysport nhập khẩu)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"사각턱보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox JETEMA (Korean)"},"cn":{"e":"方下巴肉毒素（韩国产 JETEMA）"},"vn":{"e":"Botox hàm vuông (JETEMA nội địa)"}}},{"k":"사각턱보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox Dysport (Imported)"},"cn":{"e":"方下巴肉毒素（进口 Dysport）"},"vn":{"e":"Botox hàm vuông (Dysport nhập khẩu)"}}},{"k":"풀페이스 스킨보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox JETEMA (Korean)"},"cn":{"e":"全脸皮肤肉毒素（韩国产 JETEMA）"},"vn":{"e":"Skin Botox toàn mặt (JETEMA nội địa)"}}},{"k":"풀페이스 스킨보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Dysport (Imported)"},"cn":{"e":"全脸皮肤肉毒素（进口 Dysport）"},"vn":{"e":"Skin Botox toàn mặt (Dysport nhập khẩu)"}}},{"k":"승모근/종아리보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox JETEMA (Korean)"},"cn":{"e":"斜方肌/小腿肉毒素（韩国产 JETEMA）"},"vn":{"e":"Botox cơ thang/cơ bắp chân (JETEMA nội địa)"}}},{"k":"승모근/종아리보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Dysport (Imported)"},"cn":{"e":"斜方肌/小腿肉毒素（进口 Dysport）"},"vn":{"e":"Botox cơ thang/cơ bắp chân (Dysport nhập khẩu)"}}},{"k":"다한증보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox JETEMA (Korean)"},"cn":{"e":"多汗症肉毒素（韩国产 JETEMA）"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (JETEMA nội địa)"}}},{"k":"다한증보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Dysport (Imported)"},"cn":{"e":"多汗症肉毒素（进口 Dysport）"},"vn":{"e":"Botox trị chứng tiết nhiều mồ hôi (Dysport nhập khẩu)"}}},{"k":"볼륨필러 (국산 리쥬비엘) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Rejuviel (Korean) 1cc"},"cn":{"e":"丰盈填充（韩国产 Rejuviel）1cc"},"vn":{"e":"Filler tạo phồng (Rejuviel nội địa) 1cc"}}},{"k":"입술필러 무제한 (국산 리쥬비엘)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Rejuviel (Korean)"},"cn":{"e":"嘴唇填充 不限容量（韩国产 Rejuviel）"},"vn":{"e":"Filler môi không giới hạn (Rejuviel nội địa) 1cc"}}},{"k":"눈밑 전용필러 (수입 벨로테로)","c":"필러","t":{"en":{"e":"Under-Eye Filler – Belotero (Imported)"},"cn":{"e":"眼下专用填充（进口 Belotero）"},"vn":{"e":"Filler chuyên dùng dưới mắt (Belotero nhập khẩu)"}}},{"k":"온다 1만줄","c":"리프팅","t":{"en":{"e":"ONDA 10,000 Shots"},"cn":{"e":"ONDA 1万发"},"vn":{"e":"ONDA 10.000 line"}}},{"k":"슈링크 100샷","c":"리프팅","t":{"en":{"e":"Shurink 100 Shots"},"cn":{"e":"Shurink 100发"},"vn":{"e":"Shurink 100 shot"},"tw":{"e":"Shurink 100發"}}},{"k":"볼뉴머 100샷","c":"리프팅","t":{"en":{"e":"Volnewmer 100 Shots"},"cn":{"e":"Volnewmer 100发"},"vn":{"e":"Volnewmer 100 shot"}}},{"k":"울쎄라피 프라임 100샷","c":"리프팅","t":{"en":{"e":"Ultherapy Prime 100 Shots"},"cn":{"e":"Ultherapy Prime 100发"},"vn":{"e":"Ultherapy Prime 100 shot"}}},{"k":"물광주사 2cc","c":"스킨부스터","t":{"en":{"e":"Skin Booster Injection 2cc"},"cn":{"e":"水光针 2cc"},"vn":{"e":"Tiêm căng bóng 2 cc"},"th":{"e":"ฉีดผิวฉ่ำวาว 2cc","d":""}}},{"k":"아이 리쥬란 1cc","kd":"눈가 전용","c":"스킨부스터","t":{"en":{"e":"Eye Rejuran 1cc","d":"For Eye Area Only"},"cn":{"e":"眼部 Rejuran 1cc","d":"眼周专用"},"vn":{"e":"Rejuran mắt 1cc","d":"Chuyên dùng cho vùng mắt"},"jp":{"e":"リジュランアイ1㏄","d":"目の下専用"},"th":{"e":"Eye Rejuran 1cc","d":"(สำหรับใต้ตาโดยเฉพาะ)"}}},{"k":"리쥬란힐러2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran Healer 2cc"},"cn":{"e":"Rejuran Healer 2cc"},"vn":{"e":"Rejuran Healer 2cc"}}},{"k":"쥬베룩 아이 1병","kd":"눈가 전용","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial","d":"For Eye Area Only"},"cn":{"e":"Juvelook 眼部 1瓶","d":"眼周专用"},"vn":{"e":"Juvelook Eye 1 lọ","d":"Chuyên dùng cho vùng mắt"}}},{"k":"실리프팅 1줄","kd":"팔자,턱선,심부볼 리프팅","c":"실리프팅","t":{"en":{"e":"Thread Lifting – 1 Thread","d":"Nasolabial Fold / Jawline / Deep Cheek Lifting"},"cn":{"e":"线雕 1根","d":"法令纹、下颌线、深层脸颊脂肪提升"},"vn":{"e":"Căng chỉ 1 sợi","d":"Nâng cơ rãnh cười, đường viền hàm, má hóp"},"jp":{"e":"糸リフト1本","d":"バッカルファット、ほうれい線、 顎のラインリフトアップ"},"th":{"e":"ร้อยไหม 1 เส้น","d":"(ยกกระพุ้งแก้ม ร่องแก้ม กรอบหน้า)"}}},{"k":"듀얼토닝(브라이톤+노블) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Dual Toning (Brightone + Noble) – 1 Session"},"cn":{"e":"双重嫩肤（Brightone + Noble）1次"},"vn":{"e":"Dual Toning (Brightone + Noble) 1 lần"}}},{"k":"노블 잡티레이저 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Noble Pigmentation Laser – 1 Session"},"cn":{"e":"Noble 色斑激光 1次"},"vn":{"e":"Laser đốm sắc tố Noble 1 lần"}}},{"k":"프락셀 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Fraxel – 1 Session"},"cn":{"e":"Fraxel 1次"},"vn":{"e":"Fraxel 1 lần"}}},{"k":"바디 온다 1만줄","c":"비만/다이어트","t":{"en":{"e":"Body ONDA 10,000 Shots"},"cn":{"e":"Body ONDA 1万发"},"vn":{"e":"ONDA cho body 10.000 shot"}}},{"k":"VPL 바디 지방분해주사 1회","kd":"500cc~1000c","c":"비만/다이어트","t":{"en":{"e":"VPL Body Fat-Dissolving Injection – 1 Session","d":"500cc~1000c"},"cn":{"e":"VPL 身体溶脂针 1次","d":"500cc~1000cc"},"vn":{"e":"Tiêm tan mỡ VPL Body 1 lần","d":"500cc~1000c"}}},{"k":"백옥or 신데렐라주사 1회","c":"비만/다이어트","t":{"en":{"e":"Glutathione or Cinderella IV Drip – 1 Session"},"cn":{"e":"白玉or灰姑娘针 1次"},"vn":{"e":"Tiêm Bạch Ngọc hoặc Cinderella 1 lần"}}},{"k":"이마 보톡스 기본","kd":"이마 주름 개선 자연스러운 표정 유지","c":"보톡스/윤곽주사","t":{"en":{"e":"Basic Forehead Botox","d":"Forehead Wrinkle Improvement with Natural Facial Expressions"},"cn":{"e":"额头肉毒素基本款","d":"额头皱纹改善，自然表情维持"},"th":{"e":"โบท็อกซ์หน้าผาก","d":"ช่วยลดริ้วรอยหน้าผาก พร้อมคงความเป็นธรรมชาติของสีหน้า"},"vn":{"e":"Botox trán cơ bản","d":"Cải thiện nếp nhăn trán, giữ biểu cảm tự nhiên"}}},{"k":"하이톡스 표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Hitox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"cn":{"e":"Hitox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย (เกาหลี Hitox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vn":{"e":"Hitox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"하이톡스 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Hitox Jaw Botox / Masseter Botox"},"cn":{"e":"Hitox 方下巴"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Hitox)"},"vn":{"e":"Hitox cằm vuông"}}},{"k":"하이톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Hitox Skin Botox / Full Face"},"cn":{"e":"Hitox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Hitox)"},"vn":{"e":"Hitox Skin Botox/ toàn mặt"}}},{"k":"하이톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Hitox Skin Botox / Half Face"},"cn":{"e":"Hitox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า (เกาหลี Hitox)"},"vn":{"e":"Hitox Skin Botox/ nửa mặt"}}},{"k":"하이톡스 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Hitox Skin Botox / Partial Area"},"cn":{"e":"Hitox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด (เกาหลี Hitox)"},"vn":{"e":"Hitox Skin Botox/ một phần"}}},{"k":"코어톡스  표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"cn":{"e":"Coretox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย (เกาหลี Coretox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vn":{"e":"Coretox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"코어톡스   사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Jaw Botox / Masseter Botox"},"cn":{"e":"Coretox 方下巴"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Coretox)"},"vn":{"e":"Coretox cằm vuông"}}},{"k":"코어톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Full Face"},"cn":{"e":"Coretox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Coretox)"},"vn":{"e":"Coretox Skin Botox/ toàn mặt"}}},{"k":"코어톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Half Face"},"cn":{"e":"Coretox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า (เกาหลี Coretox)"},"vn":{"e":"Coretox Skin Botox/ nửa mặt"}}},{"k":"코어톡스 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Partial Area"},"cn":{"e":"Coretox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด (เกาหลี Coretox)"},"vn":{"e":"Coretox Skin Botox/ một phần"}}},{"k":"이노톡스 표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"cn":{"e":"Innotox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย (เกาหลี Innotox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vn":{"e":"Innotox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"이노톡스 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Jaw Botox / Masseter Botox"},"cn":{"e":"Innotox 方下巴"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Innotox)"},"vn":{"e":"Innotox cằm vuông"}}},{"k":"이노톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Full Face"},"cn":{"e":"Innotox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Innotox)"},"vn":{"e":"Innotox Skin Botox/ toàn mặt"}}},{"k":"이노톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Half Face"},"cn":{"e":"Innotox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า (เกาหลี Innotox)"},"vn":{"e":"Innotox Skin Botox/ nửa mặt"}}},{"k":"이노톡스  스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Partial Area"},"cn":{"e":"Innotox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด (เกาหลี Innotox)"},"vn":{"e":"Innotox Skin Botox/ một phần"}}},{"k":"제오민 표정근육주름 1부위","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Dynamic Wrinkle Botox – 1 Area","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"cn":{"e":"Xeomin 表情肌肉皱纹 1部位","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย (นำเข้า Xeomin) 1 บริเวณ","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vn":{"e":"Xeomin cho nếp nhăn cơ biểu cảm 1 vùng","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"제오민 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Jaw Botox / Masseter Botox"},"cn":{"e":"Xeomin 方下巴"},"th":{"e":"โบท็อกซ์กราม (นำเข้า Xeomin)"},"vn":{"e":"Xeomin cằm vuông"}}},{"k":"제오민 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Full Face"},"cn":{"e":"Xeomin 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (นำเข้า Xeomin)"},"vn":{"e":"Xeomin Skin Botox/ toàn mặt"}}},{"k":"제오민 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Half Face"},"cn":{"e":"Xeomin 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า (นำเข้า Xeomin)"},"vn":{"e":"Xeomin Skin Botox/ nửa mặt"}}},{"k":"제오민 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Partial Area"},"cn":{"e":"Xeomin 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด (นำเข้า Xeomin)"},"vn":{"e":"Xeomin Skin Botox/ một phần"}}},{"k":"V-FIT 주사 (윤곽주사) 1회","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection – 1 Session"},"cn":{"e":"V-FIT 注射（轮廓针）1次"},"th":{"e":"ฉีด V-FIT (ฉีดแฟต) 1 ครั้ง"},"vn":{"e":"Tiêm V-FIT (Tiêm thon gọn line khuôn mặt) 1 lần"}}},{"k":"V-FIT 주사 (윤곽주사) 3회","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection Contouring Injection – 3 Sessions"},"cn":{"e":"V-FIT 注射（轮廓针）3次"},"th":{"e":"ฉีด V-FIT (ฉีดแฟต) 3 ครั้ง"},"vn":{"e":"Tiêm V-FIT (Tiêm thon gọn line khuôn mặt) 3 lần"}}},{"k":"슈퍼 V-FIT 주사 (슈퍼윤곽주사)1회","c":"보톡스/윤곽주사","t":{"en":{"e":"Super V-FIT Contouring Injection – 1 Session"},"cn":{"e":"超级 V-FIT 注射（超级轮廓针）1次"},"th":{"e":"ฉีด Super V-FIT (ฉีดแฟตสูตร Super) 1 ครั้ง"},"vn":{"e":"Tiêm Super V-FIT (tiêm siêu thon gọn line khuôn mặt) 1 lần"}}},{"k":"슈퍼 V-FIT 주사 (슈퍼윤곽주사)3회","c":"보톡스/윤곽주사","t":{"en":{"e":"Super V-FIT Contouring Injection Super Contouring Injection – 3 Sessions"},"cn":{"e":"超级 V-FIT 注射（超级轮廓针）3次"},"th":{"e":"ฉีด Super V-FIT (ฉีดแฟตสูตร Super) 3 ครั้ง"},"vn":{"e":"Tiêm Super V-FIT (Tiêm siêu thon gọn line khuôn mặt) 3 lần"}}},{"k":"양악주사 1회","c":"보톡스/윤곽주사","t":{"en":{"e":"Double Jaw Contouring Injection – 1 Session"},"cn":{"e":"双颌针 1次"},"th":{"e":"ฉีดหน้าเรียว 1 ครั้ง"},"vn":{"e":"Tiêm cằm 1 lần"}}},{"k":"양악주사 3회","c":"보톡스/윤곽주사","t":{"en":{"e":"Double Jaw Contouring Injection - 3 Sessions"},"cn":{"e":"双颌针 3次"},"th":{"e":"ฉีดหน้าเรียว 3 ครั้ง"},"vn":{"e":"Tiêm cằm 3 lần"}}},{"k":"아띠에르(이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Attiére Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"cn":{"e":"Attire（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Attiére ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vn":{"e":"Attiére (Trán, thái dương, má, rãnh cười, cằm lẹm) 1cc"}}},{"k":"아띠에르(코) 1CC","c":"필러","t":{"en":{"e":"Attiére Filler (Nose Filler) 1cc"},"cn":{"e":"Attire（鼻子）1CC"},"th":{"e":"Attiére (ฟิลเลอร์จมูก) 1cc"},"vn":{"e":"Attiére (Mũi ) 1CC"}}},{"k":"순수필 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Soonsoo Fill (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"cn":{"e":"Soonsoo Fill（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Soonsoo Fill ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vn":{"e":"Soonsoo Fill (Trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"순수필 (코) 1CC","c":"필러","t":{"en":{"e":"Soonsoo Fill( Nose Filler) 1"},"cn":{"e":"Soonsoo Fill（鼻子）1CC"},"th":{"e":"Soonsoo Fill (ฟิลเลอร์จมูก) 1cc"},"vn":{"e":"Soonsoo Fill (Mũi ) 1CC"}}},{"k":"아말리안 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Amalian Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"cn":{"e":"Amalian（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Amalian ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vn":{"e":"Amalian (Trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"아말리안  (코) 1CC","c":"필러","t":{"en":{"e":"Amalian Filler (Nose Filler) 1cc"},"cn":{"e":"Amalian（鼻子）1CC"},"th":{"e":"Amalian(ฟิลเลอร์จมูก) 1cc"},"vn":{"e":"Amalian (Mũi ) 1CC"}}},{"k":"와이솔루션(이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Y Solution Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"cn":{"e":"Y Solution（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Y Solution ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vn":{"e":"Y Solution (Trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"와이솔루션 (코) 1CC","c":"필러","t":{"en":{"e":"Y Solution Filler (Nose Filler) 1cc"},"cn":{"e":"Y Solution（鼻子）1CC"},"th":{"e":"Y Solution (ฟิลเลอร์จมูก) 1cc"},"vn":{"e":"Y Solution (Mũi ) 1CC"}}},{"k":"쥬비덤 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Juvederm Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"cn":{"e":"Juvederm（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Juvederm ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vn":{"e":"Juvederm (Trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"쥬비덤 (코) 1CC","c":"필러","t":{"en":{"e":"Juvederm Filler (Nose Filler) 1cc"},"cn":{"e":"Juvederm（鼻子）1CC"},"th":{"e":"Juvederm(ฟิลเลอร์จมูก) 1cc"},"vn":{"e":"Juvederm (Mũi ) 1CC"}}},{"k":"엘프요정 귀필러 1CC","c":"필러","t":{"en":{"e":"Elf Ear Filler 1cc"},"cn":{"e":"精灵耳填充 1CC"},"th":{"e":"ฟิลเลอร์หูเอลฟ์ 1cc"},"vn":{"e":"Filler tai Elf 1CC"}}},{"k":"엘프요정 귀필러 용량 무제한 귀필러(리터치포함)","c":"필러","t":{"en":{"e":"Elf Ear Filler Unlimited Ear Filler (Includes Touch-Up)"},"cn":{"e":"精灵耳填充 不限容量耳部填充（含补打）"},"th":{"e":"ฟิลเลอร์หูเอลฟ์ ฟิลเลอร์หูไม่จำกัดปริมาณ (รวมรีทัช)"},"vn":{"e":"Filler tai Elf không giới hạn dung tích (Bao gồm retouch)"}}},{"k":"셀르디엠 1병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 1 Vial"},"cn":{"e":"CellreDM 1瓶"},"th":{"e":"CellreDM 1 ขวด"},"vn":{"e":"CellreDM 1 lọ"}}},{"k":"셀르디엠 2병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 2 Vials"},"cn":{"e":"CellreDM 2瓶"},"th":{"e":"CellreDM 2 ขวด"},"vn":{"e":"CellreDM 2 lọ"}}},{"k":"셀르디엠 3병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 3 Vials"},"cn":{"e":"CellreDM 3瓶"},"th":{"e":"CellreDM 3 ขวด"},"vn":{"e":"CellreDM 3 lọ"}}},{"k":"하이엔드 안티에이징 줄기세포 스킨주사 1회","c":"줄기세포","t":{"en":{"e":"High-End Anti-Aging Stem Cell Skin Injection – 1 Session"},"cn":{"e":"高端抗衰老干细胞水光针 1次"},"th":{"e":"ฉีด High-End Anti-Aging Stem Cell Skin 1 ครั้ง"},"vn":{"e":"Tiêm skin tế bào gốc chống lão hóa High-End 1 lần"}}},{"k":"하이앤드 안티에이징 줄기세포 스킨주사 1회+ 리투오 1병","c":"줄기세포","t":{"en":{"e":"High-End Anti-Aging Stem Cell Skin Injection + 1 Vial Re2O"},"cn":{"e":"高端抗衰老干细胞水光针 1次 + Re2O 1瓶"},"th":{"e":"ฉีดสเต็มเซลล์ผิว Anti-Aging ระดับ High-End 1 ครั้ง + Re2O 1 ขวด"},"vn":{"e":"Tiêm skin tế bào gốc chống lão hóa High-End 1 lần + Re2O 1 lọ"}}},{"k":"미라셀 줄기세포 프로그램","c":"줄기세포","t":{"en":{"e":"Miracell Stem Cell Program"},"cn":{"e":"Miracell 干细胞项目"},"th":{"e":"Miracell Stem Cell Program"},"vn":{"e":"Liệu trình tế bào gốc Miracell"}}},{"k":"울쎄라 더블리프팅(턱밑울쎄라100샷+턱라인 민트파인4줄)","c":"리프팅","t":{"en":{"e":"Ultherapy Double Lifting (100 Shots Under-Chin Ultherapy + 4 Mint Fine Threads for Jawline)"},"cn":{"e":"Ultherapy 双重提升（下巴 Ultherapy 100发 + 下颌线 Mint Fine 4条）"},"th":{"e":"อัลเทอร่า Double Lifting (อัลเทอร่าใต้คาง 100 ช็อต + Mint Fine กรอบหน้า 4 เส้น)"},"vn":{"e":"Ultherapy Double Lifting (Ultherapy 100 shot vùng dưới cằm + chỉ Mint Fine 4 sợi vùng đường viền hàm)"}}},{"k":"울쎄라 300샷","c":"리프팅","t":{"en":{"e":"Ultherapy 300 Shots"},"cn":{"e":"Ultherapy 300发"},"th":{"e":"อัลเทอร่า 300 ช็อต"},"vn":{"e":"Ultherapy 300 shot"}}},{"k":"써마지 300샷","c":"리프팅","t":{"en":{"e":"Thermage 300 Shots"},"cn":{"e":"Thermage 300发"},"th":{"e":"เทอร์มาจ 300 ช็อต"},"vn":{"e":"Thermage 300 shot"},"jp":{"e":"サーマクール 300ショット"},"tw":{"e":"第四代鳳凰電波 300發"}}},{"k":"덴서티 300샷","c":"리프팅","t":{"en":{"e":"Density 300 Shots"},"cn":{"e":"Density 300发"},"th":{"e":"Density 300 ช็อต"},"vn":{"e":"Density 300 shot"}}},{"k":"덴서티 600샷","c":"리프팅","t":{"en":{"e":"Density 600 Shots"},"cn":{"e":"Density 600发"},"th":{"e":"Density 600 ช็อต"},"vn":{"e":"Density  600 shot"}}},{"k":"슈링크 300샷","c":"리프팅","t":{"en":{"e":"Shurink 300 Shots"},"cn":{"e":"Shurink 300发"},"th":{"e":"Shurink 300 ช็อต"},"vn":{"e":"Shurink 300 shot"}}},{"k":"리니어지 300샷","c":"리프팅","t":{"en":{"e":"LinearZ 300 Shots"},"cn":{"e":"LinearZ300发"},"th":{"e":"LinearZ 300 ช็อต"},"vn":{"e":"LinearZ 300 shot"}}},{"k":"인모드 fx+포마 1회","c":"리프팅","t":{"en":{"e":"InMode FX + FORMA – 1 Session"},"cn":{"e":"InMode FX+FORMA 1次"},"th":{"e":"InMode FX+FORMA 1 ครั้ง"},"vn":{"e":"InMode FX + FORMA 1 lần"}}},{"k":"리투오 1병","c":"스킨부스터","t":{"en":{"e":"Re2O – 1 Vial"},"cn":{"e":"Re2O 1瓶"},"th":{"e":"Re2O 1 ขวด"},"vn":{"e":"Re2O 1 lọ"}}},{"k":"리투오 2병","c":"스킨부스터","t":{"en":{"e":"Re2O – 2 Vials"},"cn":{"e":"Re2O 2瓶"},"th":{"e":"Re2O 2 ขวด"},"vn":{"e":"Re2O 2 lọ"}}},{"k":"리투오 3병","c":"스킨부스터","t":{"en":{"e":"Re2O – 3 Vials"},"cn":{"e":"Re2O 3瓶"},"th":{"e":"Re2O 3 ขวด"},"vn":{"e":"Re2O 3 lọ"}}},{"k":"리투오 1회+프리미엄토닝+비타민관리 10회","c":"스킨부스터","t":{"en":{"e":"Re2O 1 Session + Premium Toning + 10 Vitamin Care Sessions"},"cn":{"e":"Re2O 1次 + 高级嫩肤 + 维他命护理 10次"},"th":{"e":"Re2O 1 ครั้ง+Premium Toning+วิตามินผิว 10 ครั้ง"},"vn":{"e":"Re2O 1 lần + điều trị Toning cao cấp + chăm sóc vitamin 10 lần"}}},{"k":"리투오 1병+쥬베룩1병 (스컬트라로 변경시 20만 추가)","c":"스킨부스터","t":{"en":{"e":"Re2O 1 Vial + 1 Vial Juvelook (Upgrade to Sculptra +KRW 200,000)"},"cn":{"e":"Re2O 1瓶 + Juvelook 1瓶（更换为 Sculptra 需另加20万韩元）"},"th":{"e":"Re2O 1 ขวด+Juvelook 1 ขวด (เปลี่ยนเป็น Sculptra เพิ่ม 2 แสนวอน)"},"vn":{"e":"Re2O 1 lọ + Juvelook 1 lọ (nếu đổi sang Sculptra sẽ cộng thêm 200.000 KRW)"}}},{"k":"쥬베룩아이 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 1 Vial"},"cn":{"e":"Juvelook Eye 1瓶"},"th":{"e":"Juvelook Eye 1 ขวด"},"vn":{"e":"Juvelook mắt 1 lọ"}}},{"k":"쥬베룩아이 3병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 3 Vials"},"cn":{"e":"Juvelook Eye 3瓶"},"th":{"e":"Juvelook Eye 3 ขวด"},"vn":{"e":"Juvelook mắt 3 lọ"}}},{"k":"쥬베룩아이1병+아이리쥬란1cc+아이스킨보톡스","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial + Eye Rejuran 1cc + Ice Skin Botox"},"cn":{"e":"Juvelook Eye1瓶 + 眼部 Rejuran 1cc + 眼部肉毒素"},"th":{"e":"Juvelook Eye 1 ขวด + Eye Rejuran 1cc + สกินโบท็อกสำหรับตา"},"vn":{"e":"Juvelook mắt 1 lọ + Rejuran mắt 1cc + Ice Skin Botox"}}},{"k":"마운자로1단계 2.5mg 1box(4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 1 – 2.5mg (1 Box / 4 Pens)"},"cn":{"e":"Mounjaro 第1阶段 2.5mg 1盒（4支）"},"th":{"e":"Mounjaro Step 1 2.5mg 1 กล่อง (4 pens)"},"vn":{"e":"Mounjaro giai đoạn 1 - 2.5mg 1 hộp (4 bút tiêm)"}}},{"k":"마운자로1단계 2.5mg 3box(12pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 1 – 2.5mg (3 Boxes / 12 Pens)"},"cn":{"e":"Mounjaro 2.5mg 3盒（12支）"},"th":{"e":"Mounjaro Step 1 2.5mg 3 กล่อง (12 pens)"},"vn":{"e":"Mounjaro giai đoạn 1 - 2,5mg 3 hộp (12 bút tiêm)"}}},{"k":"마운자로2단계 5.0mg 1box(4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 2 – 5.0mg (1 Box / 4 Pens)"},"cn":{"e":"Mounjaro 第2阶段 5.0mg 1盒（4支）"},"th":{"e":"Mounjaro Step 2 5.0mg 1 กล่อง (4 pens)"},"vn":{"e":"Mounjaro giai đoạn 2 - 5.0mg 1 hộp (4 bút tiêm)"}}},{"k":"마운자로2단계 5.0mg 3box(14pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 2 –5.0mg (3 Boxes / 14 Pens)"},"cn":{"e":"Mounjaro 5.0mg 3盒（14支）"},"th":{"e":"Mounjaro Step 2 5.0mg 3 กล่อง (14 pens)"},"vn":{"e":"Mounjaro giai đoạn 2 - 5,0mg 3 hộp (14 bút tiêm)"}}},{"k":"마운자로 3단계 7.5mg 1box (4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 3 – 7.5mg (1 Box / 4 Pens)"},"cn":{"e":"Mounjaro 第3阶段 7.5mg 1盒（4支）"},"th":{"e":"Mounjaro Step 3 7.5mg 1 กล่อง (4 pens)"},"vn":{"e":"Mounjaro giai đoạn 3 - 7,5mg 1 hộp (4 bút tiêm)"}}},{"k":"여자 겨드랑이 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal (Apogee) – 1 Session"},"cn":{"e":"女性腋下脱毛（Apogee）1次"},"th":{"e":"เลเซอร์รักแร้ผู้หญิง (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông nách nữ (Apogee) 1 lần"}}},{"k":"여자 겨드랑이 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"女性腋下脱毛（Apogee）5次"},"th":{"e":"เลเซอร์รักแร้ผู้หญิง (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông nách nữ (Apogee) 5 lần"}}},{"k":"여자 팔 전체 제모(아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Full Arm Hair Removal (Apogee) – 1 Session"},"cn":{"e":"女性全臂脱毛（Apogee）1次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้หญิง (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ cánh tay nữ (Apogee) 1 lần"}}},{"k":"여자 팔 전체 제모(아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Full Arm Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"女性全臂脱毛（Apogee） 5次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้หญิง (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ cánh tay nữ (Apogee) 5 lần"}}},{"k":"여자 다리 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Full Leg Hair Removal (Apogee) – 1 Session"},"cn":{"e":"女性全腿脱毛（Apogee）1次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้หญิง (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ chân nữ (Apogee) 1 lần"}}},{"k":"여자 다리 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Full Leg Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"女性全腿脱毛（Apogee）5次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้หญิง (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ chân nữ (Apogee) 5 lần"}}},{"k":"남자 겨드랑이 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal (Apogee) – 1 Session"},"cn":{"e":"男性腋下脱毛（Apogee）1次"},"th":{"e":"เลเซอร์รักแร้ผู้ชาย (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông nách nam (Apogee) 1 lần"}}},{"k":"남자 겨드랑이 제모(아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"男性腋下脱毛（Apogee）5次"},"th":{"e":"เลเซอร์รักแร้ผู้ชาย (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông nách nam (Apogee) 5 lần"}}},{"k":"남자 팔 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Arm Hair Removal (Apogee) – 1 Session"},"cn":{"e":"男性全臂脱毛（Apogee）1次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้ชาย (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ cánh tay nam (Apogee) 1 lần"}}},{"k":"남자 팔 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Arm Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"男性全臂脱毛（Apogee）5次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้ชาย (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ cánh tay nam (Apogee) 5 lần"}}},{"k":"남자 다리 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Leg Hair Removal (Apogee) – 1 Session"},"cn":{"e":"男性全腿脱毛（Apogee）1次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้ชาย (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ chân nam (Apogee) 1 lần"}}},{"k":"남자 다리 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Leg Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"男性全腿脱毛（Apogee）5次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้ชาย (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ chân nam (Apogee) 5 lần"}}},{"k":"남자 수염 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal (Apogee) – 1 Session"},"cn":{"e":"男性全胡须脱毛（Apogee）1次"},"th":{"e":"เลเซอร์หนวดทั่วหน้าผู้ชาย (Apogee) 1 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ râu nam (Apogee) 1 lần"}}},{"k":"남자 수염 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal (Apogee) – 5 Sessions"},"cn":{"e":"男性全胡须脱毛（Apogee）5次"},"th":{"e":"เลเซอร์หนวดทั่วหน้าผู้ชาย (Apogee) 5 ครั้ง"},"vn":{"e":"Triệt lông toàn bộ râu nam (Apogee) 5 lần"}}},{"k":"큐오필 4cc","c":"필러","t":{"en":{"e":"QO Fill 4cc"},"cn":{"e":"QO Fill 4cc"},"th":{"e":"QO Fill 4cc"},"vn":{"e":"QO Fill 4cc"}}},{"k":"큐오필 8cc","c":"필러","t":{"en":{"e":"QO Fill 8cc"},"cn":{"e":"QO Fill 8cc"},"th":{"e":"QO Fill 8cc"},"vn":{"e":"QO Fill 8cc"}}},{"k":"큐오필 12cc","c":"필러","t":{"en":{"e":"QO Fill 12cc"},"cn":{"e":"QO Fill 12cc"},"th":{"e":"QO Fill 12cc"},"vn":{"e":"QO Fill 12cc"}}},{"k":"레디어스 1병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 1 Vial"},"cn":{"e":"Radiesse 1瓶"},"th":{"e":"Radiesse 1 ขวด"},"vn":{"e":"Radiesse 1 lọ"}}},{"k":"레디어스 2병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 2 Vials"},"cn":{"e":"Radiesse 2瓶"},"th":{"e":"Radiesse 2 ขวด"},"vn":{"e":"Radiesse 2 lọ"}}},{"k":"레디어스 3병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 3 Vials"},"cn":{"e":"Radiesse 3瓶"},"th":{"e":"Radiesse 3 ขวด"},"vn":{"e":"Radiesse 3 lọ"}}},{"k":"피코프락셀 (풀페이스) 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel (Full Face) – 1 Session"},"cn":{"e":"Pico Fraxel（全脸）1次"},"th":{"e":"พิโคแฟรคเซล ทั่วหน้า 1 ครั้ง"},"vn":{"e":"Pico Fraxel (Toàn mặt) 1 lần"}}},{"k":"미라젯+쥬베룩(풀페이스)1회","c":"여드름/모공/흉터","t":{"en":{"e":"MiraJet + Juvelook (Full Face) – 1 Session"},"cn":{"e":"Mirajet + Juvelook（全脸）1次"},"th":{"e":"MiraJet + Juvelook ทั่วหน้า 1 ครั้ง"},"vn":{"e":"MiraJet + Juvelook (Toàn mặt) 1 lần"}}},{"k":"포텐자(풀페이스)1회","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza (Full Face) – 1 Session"},"cn":{"e":"Potenza（全脸）1次"},"th":{"e":"โพเทนซ่า ทั่วหน้า 1 ครั้ง"},"vn":{"e":"Potenza (Toàn mặt) 1 lần"}}},{"k":"토닝5회+클라리티1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"5 Toning Sessions + 1 Clarity Session"},"cn":{"e":"嫩肤 5次 + Clarity 1次"},"th":{"e":"โทนนิ่ง 5 ครั้ง + Clarity 1 ครั้ง"},"vn":{"e":"Toning 5 lần + Clarity 1 lần"}}},{"k":"리투오1병+화이트닝토닝10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Re2O 1 Vial + 10 Whitening Toning Sessions"},"cn":{"e":"Re2O 1瓶 + 美白嫩肤10次"},"th":{"e":"Re2O 1 ขวด + โทนนิ่งผิวขาว 10 ครั้ง"},"vn":{"e":"Re2O 1 lọ + Toning làm trắng 10 lần"}}},{"k":"리투오1병+프리미엄토닝+비타민관리10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Re2O 1 Vial + Premium Toning + 10 Vitamin Care Sessions"},"cn":{"e":"Re2O 1瓶 + 高级嫩肤 + 维他命护理10次"},"th":{"e":"Re2O 1 ขวด + พรีเมียมโทนนิ่ง + วิตามินทรีตเมนต์ 10 ครั้ง"},"vn":{"e":"Re2O 1 lọ + Toning cao cấp + chăm sóc Vitamin 10 lần"}}},{"k":"리쥬란힐러 2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran Healer 2cc"},"cn":{"e":"Rejuran Healer 2cc"},"th":{"e":"Rejuran Healer 2cc"},"vn":{"e":"Rejuran Healer 2cc"}}},{"k":"리쥬란hb  2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran HB 2cc"},"cn":{"e":"Rejuran hb 2cc"},"th":{"e":"Rejuran HB 2cc"},"vn":{"e":"Rejuran HB 2cc"}}},{"k":"쥬베룩 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook – 1 Vial"},"cn":{"e":"Juvelook 1瓶"},"th":{"e":"Juvelook 1 ขวด"},"vn":{"e":"Juvelook 1 lọ"}}},{"k":"온다5만줄+1만줄","c":"리프팅","t":{"en":{"e":"ONDA 50,000 Shots + 10,000 Bonus Shots"},"cn":{"e":"ONDA 5万发 + 1万发"},"th":{"e":"ONDA 50KJ + 10KJ"},"vn":{"e":"ONDA 50.000 sợi + 10.000 sợi"}}},{"k":"온다10만줄+2만줄","c":"리프팅","t":{"en":{"e":"ONDA 100,000 Shots + 20,000 Bonus Shots"},"cn":{"e":"ONDA 10万发 + 2万发"},"th":{"e":"ONDA 100KJ + 20KJ"},"vn":{"e":"ONDA 100.000 sợi + 20.000 sợi"}}},{"k":"민트 파인 4가닥","c":"실리프팅","t":{"en":{"e":"Mint Fine Threads – 4 Threads"},"cn":{"e":"Mint Fine 4根"},"th":{"e":"Mint Fine 4 เส้น"},"vn":{"e":"Mint Fine 4 sợi"}}},{"k":"잼버 4가닥","c":"실리프팅","t":{"en":{"e":"Jamber Threads – 4 Threads"},"cn":{"e":"Jamber 4根"},"th":{"e":"Jamber 4 เส้น"},"vn":{"e":"Jamber 4 sợi"}}},{"k":"팽팽 민트 모노 100가닥","c":"실리프팅","t":{"en":{"e":"Tightening Mint Mono Threads – 100 Threads"},"cn":{"e":"紧致 Mint Mono 100根"},"th":{"e":"Mint Mono 100 เส้น"},"vn":{"e":"Chỉ Mint Mono căng 100 sợi"}}},{"k":"이마폭스 실리프팅","c":"실리프팅","t":{"en":{"e":"Forehead Fox Eye Thread Lift"},"cn":{"e":"额头 Fox 线雕"},"th":{"e":"ร้อยไหม Fox Eye หน้าผาก"},"vn":{"e":"Căng chỉ Fox vùng trán"}}},{"k":"내맘대로 커스텀 실리프팅(미미썸 시그니처)","c":"실리프팅","t":{"en":{"e":"Customized Thread Lift (MiMiSome Signature)"},"cn":{"e":"自由定制线雕（Mimisome Signature）"},"th":{"e":"ร้อยไหม Custom Design (MiMiSome Signature)"},"vn":{"e":"Căng chỉ liệu trình cá nhân theo yêu cầu (MiMiSome Signature)"}}},{"k":"골반필러 100cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 100cc"},"cn":{"e":"骨盆填充 100cc"},"th":{"e":"ฟิลเลอร์สะโพก 100cc"},"vn":{"e":"Filler vùng hông 100cc"}}},{"k":"골반필러 200cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 200cc"},"cn":{"e":"骨盆填充 200cc"},"th":{"e":"ฟิลเลอร์สะโพก 200cc"},"vn":{"e":"Filler vùng hông 200cc"}}},{"k":"골반필러 300cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 300cc"},"cn":{"e":"骨盆填充 300cc"},"th":{"e":"ฟิลเลอร์สะโพก 300cc"},"vn":{"e":"Filler vùng hông 300cc"}}},{"k":"가슴볼륨필러 100cc","c":"필러","t":{"en":{"e":"Breast Volume Filler 100cc"},"cn":{"e":"胸部丰盈填充 100cc"},"th":{"e":"ฟิลเลอร์หน้าอก 100cc"},"vn":{"e":"Filler tăng thể tích ngực 100cc"}}},{"k":"직각어깨필러","c":"필러","t":{"en":{"e":"Right-Angle Shoulder Filler"},"cn":{"e":"直角肩填充"},"th":{"e":"ฟิลเลอร์ไหล่ตรง"},"vn":{"e":"Filler tạo vai vuông"}}},{"k":"제니직각어깨필러 20cc+승모근보톡스200유닛(필러 추가시 1cc 59900원)","c":"필러","t":{"en":{"e":"Jennie Right-Angle Shoulder Filler 20cc + 200 Units+Trapezius Botox (Additional Filler KRW 59,900/cc)"},"cn":{"e":"Jennie 直角肩填充 20cc + 斜方肌肉毒素 200单位（追加填充每1cc为59900韩元）"},"th":{"e":"Jennie ฟิลเลอร์ไหล่ตรง 20cc + โบท็อกซ์บ่า 200 ยูนิต(เพิ่มฟิลเลอร์ 1cc = 59,900 วอน)"},"vn":{"e":"Filler vai vuông kiểu vai Jennie 20cc + Botox cơ thang 200 units (nếu thêm filler: 59.900 KRW / 1cc)"}}},{"k":"미니 spl 한달무제한 (인바디+비만약처방전포함)","c":"비만/다이어트","t":{"en":{"e":"Mini SPL Unlimited Monthly Package (Includes InBody Test + Obesity Medication Prescription)"},"cn":{"e":"迷你 SPL 一个月不限次数（含 InBody 检测 + 减肥药处方）"},"th":{"e":"Mini SPL ไม่จำกัด 1 เดือน (รวม Inbody + ใบสั่งยาลดน้ำหนัก)"},"vn":{"e":"Mini SPL không giới hạn 1 tháng (bao gồm đo InBody + đơn thuốc giảm cân)"}}},{"k":"지방 추출 주사 S펫 1병","c":"비만/다이어트","t":{"en":{"e":"S-FAT Fat Extraction Injection – 1 Vial"},"cn":{"e":"抽脂针 S-FAT 1瓶"},"th":{"e":"S-Fat 1 ขวด"},"vn":{"e":"Tiêm chiết xuất mỡ S-FAT 1 lọ"}}},{"k":"지방분해주사 한달무제한","c":"비만/다이어트","t":{"en":{"e":"Unlimited Fat-Dissolving Injections – 1 Month"},"cn":{"e":"溶脂针 一个月不限次数"},"th":{"e":"ฉีดสลายไขมัน ไม่จำกัด 1 เดือน"},"vn":{"e":"Tiêm tan mỡ không giới hạn 1 tháng"}}},{"k":"위고비 1개월+바디온다 1개월 (8만줄)","c":"비만/다이어트","t":{"en":{"e":"Wegovy 1-Month Program + Body ONDA 1 Month (80,000 Shots)"},"cn":{"e":"Wegovy 1个月 + Body ONDA 1个月（8万发）"},"th":{"e":"Wegovy 1 เดือน + Body ONDA 1 เดือน (80KJ)"},"vn":{"e":"Wegovy 1 tháng + Body ONDA 1 tháng (80.000 sợi)"}}},{"k":"마운자로 1개월 +바디온다 1개월(8만줄)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro 1-Month Program + Body ONDA 1 Month (80,000 Shots)"},"cn":{"e":"Mounjaro 1个月 + Body ONDA 1个月（8万发）"},"th":{"e":"Mounjaro 1 เดือน + Body ONDA 1 เดือน (80KJ)"},"vn":{"e":"Mounjaro 1 tháng + Body ONDA 1 tháng (80.000 sợi)"}}},{"k":"얼굴 점제거 1개당","kd":"0.2mm 이하 크기, 10개까지 / 첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Mole Removal – Per Lesion","d":"Up to 10 lesions under 0.2mm / First Visit Trial Price"},"vn":{"e":"Xóa nốt ruồi trên mặt 1 nốt","d":"Kích thước dưới 0,2mm, tối đa 10 cái / giá trải nghiệm lần đầu đến"}}},{"k":"흑자, 검버섯 1개","kd":"크기별 상이, 0.5cm 이내 / 첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Age Spot / Solar Lentigo Removal – Per Lesion","d":"Price Varies by Size, Within 0.5cm / First Visit Trial Price"},"vn":{"e":"Đốm sắc tố, đốm lão hóa 1 cái","d":"Giá khác nhau theo kích thước, trong phạm vi 0,5cm / giá trải nghiệm lần đầu đến"}}},{"k":"편평사마귀 / 쥐젖 / 비립종 1개당","kd":"첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Flat Wart / Skin Tag / Milia Removal – Per Lesion","d":"First Visit Trial Price"},"vn":{"e":"Mụn cóc phẳng, Skin Tag, hạt kê 1 cái","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"아쿠아필 1제, 2제","kd":"첫방문 체험가","c":"여드름/모공/흉터","t":{"en":{"e":"Aqua Peel Step 1 & 2","d":"First Visit Trial Price"},"vn":{"e":"Aqua Peel bước 1, bước 2","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"주름 보톡스 (국산)","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (Korean Brand)","d":"First Visit Trial Price"},"vn":{"e":"Botox xóa nhăn (nội địa)","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"사각턱 보톡스 (국산)","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox (Korean Brand)","d":"First Visit Trial Price"},"vn":{"e":"Botox gọn hàm vuông (nội địa)","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"브이핏 윤곽주사 1cc","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection 1cc","d":"First Visit Trial Price"},"vn":{"e":"Tiêm V-FIT tạo đường khuôn mặt 1cc","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"프리미엄 명품 피코토닝 1회","kd":"첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Premium Pico Toning – 1 Session","d":"First Visit Trial Price"},"vn":{"e":"Pico Toning cao cấp 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"여성 아포지 인중 or 겨드랑이 제모 1회","kd":"첫방문 체험가","c":"제모","t":{"en":{"e":"Women’s Apogee Upper Lip or Underarm Hair Removal – 1 Session","d":"First Visit Trial Price"},"vn":{"e":"Triệt lông Apogee nữ vùng ria mép hoặc nách 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"남성 아포지 인중 제모 1회","kd":"첫방문 체험가","c":"제모","t":{"en":{"e":"Men’s Apogee Upper Lip Hair Removal – 1 Session","d":"First Visit Trial Price"},"vn":{"e":"Triệt lông Apogee nam vùng ria mép 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"표정근육주름 1부위","kd":"하이톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Hutox (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vn":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"HiTox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"코어톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Coretox  (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vn":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"CoreTox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"이노톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Innotox  (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vn":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"Innotox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"제오민 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Xeomin (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vn":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"Xeomin (Chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"사각턱 보톡스","kd":"하이톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Hutox"},"vn":{"e":"Botox gọn cằm vuông","d":"HiTox"}}},{"k":"사각턱 보톡스","kd":"코어톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Coretox"},"vn":{"e":"Botox gọn cằm vuông","d":"CoreTox"}}},{"k":"사각턱 보톡스","kd":"이노톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Innotox"},"vn":{"e":"Botox gọn cằm vuông","d":"Innotox"}}},{"k":"사각턱 보톡스","kd":"제오민","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Xeomin"},"vn":{"e":"Botox gọn cằm vuông","d":"Xeomin"}}},{"k":"스킨보톡스","kd":"하이톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Hutox"},"vn":{"e":"Skin Botox","d":"HiTox"}}},{"k":"스킨보톡스","kd":"코어톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Coretox"},"vn":{"e":"Skin Botox","d":"CoreTox"}}},{"k":"스킨보톡스","kd":"이노톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Innotox"},"vn":{"e":"Skin Botox","d":"Innotox"}}},{"k":"스킨보톡스","kd":"제오민","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Xeomin"},"vn":{"e":"Skin Botox","d":"Xeomin"}}},{"k":"브이핏 윤곽주사 12cc","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection 12cc"},"vn":{"e":"Tiêm V-FIT tạo đường khuôn mặt 12cc"}}},{"k":"미미썸 양악주사","c":"보톡스/윤곽주사","t":{"en":{"e":"MiMiSome Double Jaw Contouring Injection"},"vn":{"e":"Tiêm tạo gọn hàm MiMiSome"}}},{"k":"코 조각주사","c":"보톡스/윤곽주사","t":{"en":{"e":"Nose Contouring Injection"},"vn":{"e":"Tiêm tạo dáng mũi"}}},{"k":"볼륨 필러 1cc","kd":"아띠에르","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Attiére"},"vn":{"e":"Filler tạo volume 1cc","d":"Atiere"}}},{"k":"볼륨 필러 1cc","kd":"아말리안","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Amalian"},"vn":{"e":"Filler tạo volume 1cc","d":"Amalian"}}},{"k":"볼륨 필러 1cc","kd":"쥬비덤","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Juvederm"},"vn":{"e":"Filler tạo volume 1cc","d":"Juvederm"}}},{"k":"벨로테로 소프트 1cc","kd":"눈두덩 & 눈밑고랑 꺼짐 개선","c":"필러","t":{"en":{"e":"Belotero Soft 1cc","d":"Upper Eyelid & Tear Trough Volume Correction"},"vn":{"e":"Belotero Soft 1cc","d":"Cải thiện vùng mí mắt & rãnh dưới mắt bị hõm"}}},{"k":"레스틸렌 키스 1cc","c":"필러","t":{"en":{"e":"Restylane Kysse 1cc"},"vn":{"e":"Restylane Kysse 1cc"}}},{"k":"큐오필 줄기세포","c":"필러","t":{"en":{"e":"QO Fill Stem Cell Filler"},"vn":{"e":"QO Fill tế bào gốc"}}},{"k":"인모드 FX 1부위","c":"리프팅","t":{"en":{"e":"InMode FX – 1 Area"},"vn":{"e":"InMode FX 1 vùng"}}},{"k":"인모드 FORMA","c":"리프팅","t":{"en":{"e":"InMode FORMA"},"vn":{"e":"InMode FORMA"}}},{"k":"온다 3만줄 + LDM 초음파 리프팅 12분","c":"리프팅","t":{"en":{"e":"ONDA 30,000 Shots + LDM Ultrasound Lifting 12 min"},"vn":{"e":"ONDA 30.000 lines + LDM nâng cơ bằng sóng siêu âm 12 phút"}}},{"k":"물광주사 6cc","c":"스킨부스터","t":{"en":{"e":"Skin Booster Injection 6cc"},"vn":{"e":"Tiêm căng bóng 6cc"}}},{"k":"리쥬란 2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran 2cc"},"vn":{"e":"Rejuran 2cc"}}},{"k":"쥬베룩 스킨 4cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Skin 4cc"},"vn":{"e":"Juvelook Skin 4cc"}}},{"k":"쥬베룩 볼륨 2cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Volume 2cc"},"vn":{"e":"Juvelook Volume 2cc"}}},{"k":"쥬베룩 아이 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 1 Vial"},"vn":{"e":"Juvelook mắt 1 lọ"},"jp":{"e":"ジュベルックアイ1バイアル","d":"目の下専用"},"cn":{"e":"Juvelook 眼部 1瓶","d":"眼周专用"},"th":{"e":"Juvelook Eye 1 ขวด","d":"(สำหรับใต้ตาโดยเฉพาะ)"}},"kd":"눈밑 전용"},{"k":"미라셀 스킨주사 30cc","c":"줄기세포","t":{"en":{"e":"Miracell Skin Injection 30cc"},"vn":{"e":"Tiêm skin Miracell 30cc"}}},{"k":"미라셀 정맥주사 120cc","c":"줄기세포","t":{"en":{"e":"Miracell IV Therapy 120cc"},"vn":{"e":"Miracell truyền tĩnh mạch 120cc"}}},{"k":"민트실 1줄","c":"실리프팅","t":{"en":{"e":"Mint Thread – 1 Thread"},"vn":{"e":"Chỉ Mint 1 sợi"}}},{"k":"민트실 4줄","c":"실리프팅","t":{"en":{"e":"Mint Thread – 4 Threads"},"vn":{"e":"Chỉ Mint 4 sợi"}}},{"k":"눈썹 / 이마 거상 실리프팅 4줄","c":"실리프팅","t":{"en":{"e":"Brow / Forehead Lift Thread Lifting – 4 Threads"},"vn":{"e":"Căng chỉ mày / trán 4 sợi"}}},{"k":"하이코 실리프팅 2줄","c":"실리프팅","t":{"en":{"e":"Hiko Nose Thread Lift – 2 Threads"},"vn":{"e":"Căng chỉ Hiko 2 sợi"}}},{"k":"브라이톤 토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Brightone Toning – 1 Session"},"vn":{"e":"Brightone Toning 1 lần"}}},{"k":"포텐자 (모공 / 기미 / 여드름) + 연어 재생 앰플 관리 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Potenza (Pores / Melasma / Acne) + Salmon Regeneration Ampoule Care – 1 Session"},"vn":{"e":"Potenza (lỗ chân lông / nám / mụn) + chăm sóc ampoule tái tạo từ cá hồi 1 lần"}}},{"k":"엑셀브이 플러스 제네시스 + 진정 앰플 관리 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Excel V Plus Genesis + Soothing Ampoule Care – 1 Session"},"vn":{"e":"Excel V Plus Genesis + chăm sóc Ampoule làm dịu da 1 lần"}}},{"k":"헐리우드 스펙트라 프리미엄 토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Premium Toning – 1 Session"},"vn":{"e":"Hollywood Spectra Premium Toning 1 lần"}}},{"k":"주근깨, 잡티 레이저 1회","kd":"클라리티 Pro","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Freckle & Pigmentation Laser – 1 Session","d":"Clarity Pro"},"vn":{"e":"Laser trị tàn nhang, đốm sắc tố 1 lần","d":"Clarity Pro"}}},{"k":"브라이톤 듀얼토닝 10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Brightone Dual Toning – 10 Sessions"},"vn":{"e":"Brightone Dual Toning 10 lần"}}},{"k":"헐리우드 스펙트라 프리미엄 토닝 12회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Premium Toning – 12 Sessions"},"vn":{"e":"Hollywood Spectra Premium Toning 12 lần"}}},{"k":"아쿠아필 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Aqua Peel – 1 Session"},"vn":{"e":"Aqua Peel 1 lần"}}},{"k":"라라필 1회","c":"여드름/모공/흉터","t":{"en":{"e":"LHALA Peel – 1 Session"},"vn":{"e":"LHALA Peel 1 lần"}}},{"k":"LDM 재생관리 1회","c":"여드름/모공/흉터","t":{"en":{"e":"LDM Regeneration Care – 1 Session"},"vn":{"e":"Chăm sóc tái tạo LDM 1 lần"}}},{"k":"여드름관리 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Acne Care Treatment – 1 Session"},"vn":{"e":"Chăm sóc điều trị mụn 1 lần"}}},{"k":"피코프락셀 1회","kd":"부분","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel – 1 Session","d":"Partial Area"},"vn":{"e":"Pico Fraxel 1 lần","d":"1 phần"}}},{"k":"여자 인중 제모 5회","c":"제모","t":{"en":{"e":"Women’s Upper Lip Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông ria mép nữ 5 lần"}}},{"k":"여자 겨드랑이 제모 5회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông nách 5 lần"}}},{"k":"여자 비키니라인 제모 5회","c":"제모","t":{"en":{"e":"Women’s Bikini Line Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông vùng bikini 5 lần"}}},{"k":"여자 브라질리언 제모 5회","c":"제모","t":{"en":{"e":"Women’s Brazilian Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông toàn vùng kín 5 lần"}}},{"k":"남자 인중 제모 5회","c":"제모","t":{"en":{"e":"Men’s Upper Lip Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông ria mép nam 5 lần"}}},{"k":"남자 하관전체 제모 5회","c":"제모","t":{"en":{"e":"Men’s Full Lower Face Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông toàn bộ vùng hàm dưới 5 lần"}}},{"k":"남자 겨드랑이 제모 5회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal – 5 Sessions"},"vn":{"e":"Triệt lông nách 5 lần"}}},{"k":"얼굴 전체 솜털 제모 1회","c":"제모","t":{"en":{"e":"Full Face Fine Hair Removal – 1 Session"},"vn":{"e":"Triệt lông tơ toàn mặt 1 lần"}}},{"k":"퀵영양수액 1회","kd":"태반 / 비타민/ 백옥 / 신데렐라 중 선택","c":"비만/다이어트","t":{"en":{"e":"Quick Nutrition IV Drip – 1 Session","d":"Choose 1: Placenta / Vitamin / Glutathione / Cinderella IV"},"vn":{"e":"Truyền dinh dưỡng nhanh 1 lần"}}},{"k":"퀵영양수액 5회","kd":"태반 / 비타민/ 백옥 / 신데렐라 중 선택","c":"비만/다이어트","t":{"en":{"e":"Quick Nutrition IV Drip – 5 Sessions"},"vn":{"e":"Truyền dinh dưỡng nhanh 5 lần","d":"Chọn 1 trong: Nhau thai / Vitamin / Bạch Ngọc / Cinderella"}}},{"k":"고농도 메가비타민 1회","kd":"고농도 비타민C","c":"비만/다이어트","t":{"en":{"e":"High-Dose Mega Vitamin IV – 1 Session","d":"High-Dose Vitamin C"},"vn":{"e":"Mega Vitamin nồng độ cao 1 lần"}}},{"k":"고농도 메가비타민 5회","kd":"고농도 비타민C","c":"비만/다이어트","t":{"en":{"e":"High-Dose Mega Vitamin IV – 5 Sessions"},"vn":{"e":"Mega Vitamin nồng độ cao 5 lần","d":"Vitamin C nồng độ cao"}}},{"k":"항노화주사 1회","kd":"포도당 + 비타민C + 비타민B","c":"비만/다이어트","t":{"en":{"e":"Anti-Aging Injection – 1 Session","d":"Glucose + Vitamin C + Vitamin B"},"vn":{"e":"Tiêm chống lão hóa 1 lần"},"jp":{"e":"抗老化注射1回","d":"疲労回復、抗酸化"},"cn":{"e":"抗衰老针 1次","d":"疲劳恢复、抗氧化"},"th":{"e":"ฉีดชะลอวัย 1 ครั้ง","d":"ฟื้นฟูร่างกาย, ต้านอนุมูลอิสระ"}}},{"k":"항노화주사 5회","kd":"포도당 + 비타민C + 비타민B","c":"비만/다이어트","t":{"en":{"e":"Anti-Aging Injection – 5 Sessions"},"vn":{"e":"Tiêm chống lão hóa 5 lần"}}},{"k":"신비주사 1회","kd":"포도당 + 비타민C + 신데렐라","c":"비만/다이어트","t":{"en":{"e":"Shinbi IV Drip – 1 Session"},"vn":{"e":"Tiêm Shinbi 1 lần"}}},{"k":"신비주사 5회","kd":"포도당 + 비타민C + 신데렐라","c":"비만/다이어트","t":{"en":{"e":"Shinbi IV Drip – 5 Sessions"},"vn":{"e":"Tiêm Shinbi 5 lần","d":"Glucose + Vitamin C + Vitamin B"}}},{"k":"오로라주사 1회","kd":"포도당 + 비타민C + 백옥","c":"비만/다이어트","t":{"en":{"e":"Aurora IV Drip – 1 Session","d":"Glucose + Vitamin C + Glutathione"},"vn":{"e":"Tiêm Aurora 1 lần"}}},{"k":"오로라주사 5회","kd":"포도당 + 비타민C + 백옥","c":"비만/다이어트","t":{"en":{"e":"Aurora IV Drip – 5 Sessions"},"vn":{"e":"Tiêm Aurora 5 lần","d":"Glucose + Vitamin C + Bạch Ngọc"}}},{"k":"소녀주사 1회","kd":"고농도 비타민C + 신데렐라 + 백옥 + 태반","c":"비만/다이어트","t":{"en":{"e":"Girl IV Drip – 1 Session","d":"High-Dose Vitamin C + Cinderella + Glutathione + Placenta"},"vn":{"e":"Tiêm tiêm trẻ hóa 1 lần"}}},{"k":"소녀주사 5회","kd":"고농도 비타민C + 신데렐라 + 백옥 + 태반","c":"비만/다이어트","t":{"en":{"e":"Girl IV Drip – 5 Sessions"},"vn":{"e":"Tiêm tiêm trẻ hóa 5 lần","d":"Vitamin C nồng độ cao + Cinderella + Bạch Ngọc + nhau thai"}}},{"k":"와인주사 추가","kd":"모든 수액에 추가 가능","c":"비만/다이어트","t":{"en":{"e":"Wine Injection Add-On","d":"Can Be Added to Any IV Drip"},"vn":{"e":"Thêm tiêm Wine Injection","d":"Có thể thêm vào tất cả loại truyền dịch"}}},{"k":"무통증 수면 마취","kd":"장시간 풀수면 마취","c":"성형수술","t":{"en":{"e":"Painless Sedation Anesthesia","d":"Long-Duration Full Sedation Anesthesia"},"vn":{"e":"Gây mê ngủ không đau","d":"Gây mê ngủ sâu thời gian dài"}}},{"k":"저통증 가스 마취","kd":"에어녹스 스마일 가스 마취","c":"성형수술","t":{"en":{"e":"Low-Pain Gas Anesthesia","d":"AirNox Smile Gas Sedation"},"vn":{"e":"Gây mê khí giảm đau mức thấp","d":"Gây mê khí AirNox Smile Gas"}}},{"k":"보조개 한쪽","kd":"긴보조개 디자인 가능","c":"성형수술","t":{"en":{"e":"Dimple Creation – One Side","d":"Long Dimple Design Available"},"vn":{"e":"Tạo lúm đồng tiền 1 bên","d":"Có thể thiết kế lúm đồng tiền dài"},"jp":{"e":"えくぼ形成片側","d":""},"cn":{"e":"酒窝（单侧）","d":""},"th":{"e":"ทำลักยิ้ม (1 ข้าง)","d":""}}},{"k":"3D 눈밑 돌출 지방 제거 + 눈밑 & 앞볼 지방 이식","c":"성형수술","t":{"en":{"e":"3D Under-Eye Fat Removal + Under-Eye & Front Cheek Fat Grafting"},"vn":{"e":"Loại bỏ mỡ phồng dưới mắt 3D + cấy mỡ dưới mắt & vùng má trước"}}},{"k":"360도 무제한 지방 흡입","kd":"팔전체 / 상 or 하복부 / 허벅지 중 선택","c":"성형수술","t":{"en":{"e":"360° Unlimited Liposuction","d":"Choose 1: Full Arms / Upper or Lower Abdomen / Thighs"},"vn":{"e":"Hút mỡ 360 độ không giới hạn","d":"Chọn: toàn bộ cánh tay / bụng trên hoặc bụng dưới / đùi"}}},{"k":"미미썸 한정","t":{"en":{"e":"Mimisome Limited"},"jp":{"e":"ミミソム限定"},"tw":{"e":"MIMISOME 限定"}}},{"k":"메조주사 + LED","t":{"en":{"e":"Mesotherapy Injection + LED"},"jp":{"e":"メソセラピー注射 + LED"},"tw":{"e":"頭皮中胚層療法注射 + LED + 藥物處方"}}},{"k":"메조주사 + LED + 약물처방","t":{"en":{"e":"Mesotherapy Injection + LED + Medication Prescription"},"jp":{"e":"メソセラピー注射 + LED + 薬処方"},"tw":{"e":"頭皮中胚層療法注射 + LED + 藥物處方"}}},{"k":"*VAT포함","t":{"en":{"e":"VAT included"},"jp":{"e":"VAT込み"},"tw":{"e":"含增值稅"}}},{"k":"휑한 가르마 꽉! 잡는","t":{"en":{"e":"Targets thinning hair part lines"},"jp":{"e":"分け目の薄さをしっかりカバーする"},"tw":{"e":"稀疏分線全面改善"}}},{"k":"3-Step 메디컬 탈모·두피 재생 솔루션","t":{"en":{"e":"A 3 step medical hair loss and scalp regeneration solution"},"jp":{"e":"3-Step 医療発毛・頭皮再生ソリューション"},"tw":{"e":"3-Step 醫療級落髮頭皮再生方案"}}},{"k":"주사하고, 빛으로 채우는 빈틈없는 모근 강화 프로그램","t":{"en":{"e":"Fill from within by taking medication, injecting, and using light. Pore strengthening program"},"jp":{"e":"注入・光で補う隙のない毛根強化プログラム"},"tw":{"e":"透過口服、注射,再以光能填補空隙的完整毛根強化計畫"}}},{"k":"이런 분들께 강력 추천합니다!","t":{"en":{"e":"Recommended for these people"},"jp":{"e":"次のような方に強くお勧めします!"},"tw":{"e":"強烈推薦給以下族群!"}}},{"k":"하루가 다르게 머리카락이 얇아지고 힘이 없는 분","t":{"en":{"e":"Those whose hair part is widening day by day and showing no improvement"},"jp":{"e":"日々髪が細くなり、ハリ・コシがなくなってきた方"},"tw":{"e":"頭髮一天比一天變細、缺乏彈力者"}}},{"k":"탈모 샴푸나 영양제만으로는 한계를 느끼시는 분","t":{"en":{"e":"Those who feel limits with only hair loss shampoo and nutrition"},"jp":{"e":"育毛シャンプーやサプリだけでは限界を感じている方"},"tw":{"e":"僅靠防脫洗髮精或營養補充品感到效果有限者"}}},{"k":"스트레스, 다이어트, 산후 등으로 두피 환경이 무너지신 분","t":{"en":{"e":"Those whose scalp environment has collapsed due to stress, dieting, or postpartum changes"},"jp":{"e":"ストレス・ダイエット・産後などで頭皮環境が乱れている方"},"tw":{"e":"因壓力、減肥、產後等導致頭皮環境受損者"}}},{"k":"단일 치료가 아닌, 체계적인 메디컬 관리가 필요하신 분","t":{"en":{"e":"Those who need systematic medical care rather than single treatments"},"jp":{"e":"単一治療ではなく、体系的な医療ケアが必要な方"},"tw":{"e":"需要系統化醫療管理,而非單一療程者"}}},{"k":"3-Step 트리플 시너지, 왜 다를까요?","t":{"en":{"e":"3-Step Why is the triple synergy different"},"jp":{"e":"3-Step トリプルシナジー、なぜ違うのか?"},"tw":{"e":"3-Step 三重協同作用,為何不同?"}}},{"k":"단순히 겉만 관리하는 것이 아닙니다. 탈모의 원인 차단부터 두피 속 섬유아세포를 자극하는 재생 단계까지, 의학적 근거를 바탕으로 모근의 힘을 길러줍니다.","t":{"en":{"e":"Not a simple management. It works precisely on each stage from the root cause of hair loss to the scalp environment. Based on medical evidence, it strengthens the power of hair roots"},"jp":{"e":"表面的なケアだけではありません。脱毛の原因を抑制する段階から、頭皮内部の線維芽細胞を刺激する再生段階まで、医学的根拠に基づき毛根の力を高めます。"},"tw":{"e":"不只是表面護理。從阻斷落髮根本原因到刺激頭皮內纖維母細胞的再生階段,以醫學根據強化毛根力量"}}},{"k":"STEP 1 [차단] 메디컬 약물요법","t":{"en":{"e":"STEP 1 Medical medication therapy"},"jp":{"e":"STEP 1 【抑制】医療薬物療法"},"tw":{"e":"STEP 1 【阻斷】醫療藥物療法"}}},{"k":"탈모를 유발하는 근본적인 원인을 의학적으로 억제하여, 더 이상의 진행을 막아주는 가장 탄탄한 기초 단계입니다.","t":{"en":{"e":"Medically suppresses the root causes that induce hair loss. Becomes a solid foundation stage that supports further progression"},"jp":{"e":"脱毛の根本原因を医学的に抑制し、進行を防ぐ最も重要な基礎段階です。"},"tw":{"e":"以醫學方式抑制造成落髮的根本原因,阻止進一步惡化,是最穩固的基礎階段"}}},{"k":"STEP 2 [재생] 두피 메조테라피 주사","t":{"en":{"e":"STEP 2 Regeneration scalp mesotherapy injection"},"jp":{"e":"STEP 2 【再生】頭皮メソセラピー注射"},"tw":{"e":"STEP 2 【再生】頭皮中胚層療法注射"}}},{"k":"모낭 생존에 필수적인 유효 성분(영양 물질)을 모근 가까이 직접 주사합니다. 두피 속부터 건강한 세포 재생을 유도하여 튼튼한 모발이 자랄 수 있는 토대를 만듭니다.","t":{"en":{"e":"Directly supplies essential active ingredients needed for hair growth. Rebuilds the scalp environment from the scalp. Helps hair to grow thicker"},"jp":{"e":"毛包の生存に必要な有効成分(栄養成分)を頭皮の深部に直接注入します。頭皮内部から促進健康な細胞再生から、強い毛髪が育つ土台を作ります。"},"tw":{"e":"將毛囊存活所需的有效成分(營養物質)直接注入頭皮深層。從頭皮內部促進健康細胞再生,打造強韌毛髮生長基礎"}}},{"k":"STEP 3 [활성] LED 광에너지 테라피","t":{"en":{"e":"STEP 3 Activation LED care therapy"},"jp":{"e":"STEP 3 【活性】LED光エネルギー療法"},"tw":{"e":"STEP 3 【活化】LED光能療法"}}},{"k":"특수 파장의 빛이 두피 깊숙이 침투하여 미세 혈액순환을 촉진합니다. 메조테라피 성분의 흡수율을 극대화하고, 두피 섬유아세포를 자극해 재생 시너지를 한층 더 높여줍니다.","t":{"en":{"e":"Special wavelength light penetrates deep into the scalp. Promotes regeneration of micro blood vessels and hair follicles. Maximizes mesotherapy effects and increases hair root cell regeneration activity"},"jp":{"e":"特殊な波長の光が頭皮の奥深くまで浸透し、微細な血行を促進します。メソセラピー成分の吸収を最大化し、線維芽細胞の活性化と再生シナジーをさらに高めます。"},"tw":{"e":"特殊波長光深入頭皮,促進微循環,最大化中胚層療法成分吸收,提升頭皮纖維母細胞刺激與再生協同效果"}}},{"k":"시술 시간","t":{"en":{"e":"Procedure time"},"jp":{"e":"施術時間"},"tw":{"e":"療程時間"}}},{"k":"약 30~40분 (빠르고 간편하게!)","t":{"en":{"e":"about 30 to 40 minutes / Fast and simple"},"jp":{"e":"約30~40分 (スピーディーで簡単)"},"tw":{"e":"約30~40分鐘 (快速且便利)"}}},{"k":"마취 여부","t":{"en":{"e":"Anesthesia"},"jp":{"e":"麻酔の有無"},"tw":{"e":"是否麻醉"}}},{"k":"불필요 (또는 연고 마취 진행)","t":{"en":{"e":"Not required / Topical anesthesia may be used"},"jp":{"e":"不要 (または麻酔クリーム使用)"},"tw":{"e":"不需要 (或可進行表面麻醉)"}}},{"k":"일상 생활","t":{"en":{"e":"Daily life"},"jp":{"e":"日常生活"},"tw":{"e":"日常生活"}}},{"k":"시술 직후 바로 일상생활 복귀 가능","t":{"en":{"e":"Return to daily activities possible immediately after the procedure"},"jp":{"e":"施術直後から日常生活に復帰可能"},"tw":{"e":"療程後可立即恢復日常生活"}}},{"k":"추천 주기","t":{"en":{"e":"Recommended cycle"},"jp":{"e":"推奨周期"},"tw":{"e":"建議週期"}}},{"k":"1~2주 간격 (개인별 두피 상태에 따라 맞춤 상담 진행)","t":{"en":{"e":"Every 1 to 2 weeks / Customized depending on individual scalp condition"},"jp":{"e":"1~2週間間隔 (頭皮状態に応じて個別カウンセリング)"},"tw":{"e":"每1~2週一次 (依個人頭皮狀況進行客製化諮詢)"}}},{"k":"시크릿 이너볼륨 필러","t":{"en":{"e":"Secret Inner Volume Filler"},"jp":{"e":"シークレットインナーボリュームフィラー"},"cn":{"e":"秘密 内层 丰盈填充剂"},"tw":{"e":"秘密 內層 豐盈填充劑"},"th":{"e":"ฟิลเลอร์ Secret Inner Volume"}},"c":"필러"},{"k":"5cc 715,000원","t":{"en":{"e":"5cc 715,000 KRW"},"jp":{"e":"5cc 715,000 KRW"},"cn":{"e":"5cc 715,000 KRW"},"tw":{"e":"5cc 715,000 KRW"},"th":{"e":"5cc 715,000 KRW"}}},{"k":"속부터 차오르는 자신감","t":{"en":{"e":"Confidence that fills from within"},"jp":{"e":"内側から変わる、自信"},"cn":{"e":"从内在开始变化的自信感秘密"},"tw":{"e":"從內在開始改變的自信感秘密"},"th":{"e":"ความมั่นใจที่เต็มเต็มจากภายใน"}}},{"k":"여성 인티메이트 전용 시술","t":{"en":{"e":"Intimate treatment for women"},"jp":{"e":"女性デリケートゾーン専用施術"},"cn":{"e":"女性私密部位专用施术"},"tw":{"e":"女性私密部位專用療程"},"th":{"e":"หัตถการเฉพาะจุดซ่อนเร้นสำหรับผู้หญิง"}}},{"k":"만족도 UP!","t":{"en":{"e":"VOLUME UP!"},"jp":{"e":"ボリュームUP!"},"cn":{"e":"饱满UP!"},"tw":{"e":"飽滿UP!"},"th":{"e":"ฉันเพิ่มได้!"}}},{"k":"자신감 UP!","t":{"en":{"e":"CONFIDENCE UP!"},"jp":{"e":"自信UP!"},"cn":{"e":"自信UP!"},"tw":{"e":"自信UP!"},"th":{"e":"เพิ่มความมั่นใจ!"}}},{"k":"시크릿 이너볼륨 필러란?","t":{"en":{"e":"What is Secret Inner Volume Filler?"},"jp":{"e":"シークレットインナーボリュームフィラーとは?"},"cn":{"e":"什么是私密内在丰盈填充?"},"tw":{"e":"秘密內在豐盈填充是什麼?"},"th":{"e":"ฟิลเลอร์เพิ่มวอลลุ่มภายในจุดซ่อนเร้นคืออะไร?"}}},{"k":"시크릿 이너볼륨 필러는 인체 친화적인 필러 성분을 사용하여 질 등의 여성 인티메이트 부위의 볼륨과 탄력을 자연스럽게 개선하는 시술입니다","t":{"en":{"e":"Secret Inner Volume Filler is a treatment that uses biocompatible filler ingredients to naturally improve volume and elasticity in women's intimate areas such as the vaginal region."},"jp":{"e":"シークレットインナーボリュームフィラーは、人体に優しいフィラー成分を使用し、膣などの女性のデリケートゾーンのボリュームと弾力を自然に改善する施術です"},"cn":{"e":"私密内在丰盈填充是一种采用对人体友好的填充成分,针对阴道等女性私密部位,自然改善其丰盈度与弹性的治疗方式。"},"tw":{"e":"秘密內在豐盈填充 採用對人體友善的填充成分 針對陰道等女性私密部位 自然改善其豐盈度與彈性的療程"},"th":{"e":"ฟิลเลอร์เพิ่มวอลลุ่มภายในจุดซ่อนเร้น เป็นหัตถการที่ใช้ฟิลเลอร์ที่เป็นมิตรต่อร่างกาย เพื่อช่วยปรับวอลลุ่มและความยืดหยุ่นของจุดซ่อนเร้นของผู้หญิง เช่น บริเวณช่องคลอด ให้ดีขึ้นอย่างเป็นธรรมชาติ"}}},{"k":"빠른 효과로 눈에 보이고 몸으로 느껴지는 변화","t":{"en":{"e":"Fast results you can see with your eyes and feel with your body"},"jp":{"e":"即効性があり 目で見て分かり 身体で実感できる変化"},"cn":{"e":"快速见效 不仅看得见 身体也能感受到的变化"},"tw":{"e":"快速見效 不只看得見 身體也能實際感受到的改變"},"th":{"e":"เห็นผลรวดเร็ว สังเกตได้ด้วยตา และสัมผัสได้จริง"}}},{"k":"확실한 탄력","t":{"en":{"e":"Firm elasticity"},"jp":{"e":"しっかりとした弾力"},"cn":{"e":"明显弹性"},"tw":{"e":"明顯彈性"},"th":{"e":"ความกระชับชัดเจน"}}},{"k":"자연스러운 볼륨감","t":{"en":{"e":"Natural volume"},"jp":{"e":"自然なボリューム感"},"cn":{"e":"自然饱满感"},"tw":{"e":"自然豐盈感"},"th":{"e":"วอลลุ่มดูเป็นธรรมชาติ"}}},{"k":"성감 극대화","t":{"en":{"e":"Enhanced confidence"},"jp":{"e":"自信の向上"},"cn":{"e":"自信提升"},"tw":{"e":"自信提升"},"th":{"e":"เพิ่มความพึงพอใจในตัวเอง"}}},{"k":"미미썸 시크릿 볼륨 필러 추천 대상","t":{"en":{"e":"Secret Volume Filler Recommended for"},"jp":{"e":"ミミサム シークレットインナーボリュームフィラー おすすめの方"},"cn":{"e":"秘密内在丰盈填充 推荐对象"},"tw":{"e":"秘密內在豐盈填充 推薦對象"},"th":{"e":"ผู้ที่เหมาะกับฟิลเลอร์ Secret Volume ของ Mimisome"}}},{"k":"성감을 보호하거나 개선하고 싶으신 분","t":{"en":{"e":"Those who want to enhance or improve sensitivity"},"jp":{"e":"出産や加齢により弾力の低下を感じる方"},"cn":{"e":"因分娩或老化导致弹性下降的人"},"tw":{"e":"因生產或老化導致彈性下降者"},"th":{"e":"ผู้ที่ต้องการเสริมหรือปรับปรุงความรู้สึกขณะสัมผัส"}}},{"k":"파트너의 만족도를 빠르게 높이고 싶으신 분","t":{"en":{"e":"Those seeking quick volume satisfaction"},"jp":{"e":"デリケートゾーンのボリューム感を改善したい方"},"cn":{"e":"希望改善私密部位饱满感的人"},"tw":{"e":"希望改善私密部位豐盈感者"},"th":{"e":"ผู้ที่ต้องการเพิ่ม Inner Volume อย่างรวดเร็ว"}}},{"k":"자연스럽게 자신감을 회복하고 싶으신 분","t":{"en":{"e":"Those who want to regain confidence naturally"},"jp":{"e":"自然に自信を取り戻したい方"},"cn":{"e":"希望自然恢复自信感的人"},"tw":{"e":"希望自然恢復自信感者"},"th":{"e":"ผู้ที่ต้องการฟื้นคืนความมั่นใจให้เป็นธรรมชาติ"}}},{"k":"섬세한 맞춤 디자인","t":{"en":{"e":"Delicate Customized Design"},"jp":{"e":"繊細な触感の変化"},"cn":{"e":"细腻的触感改变"},"tw":{"e":"細膩觸感的改變"},"th":{"e":"การออกแบบหัตถการอย่างละเอียด"}}},{"k":"안전성 + 지속력","t":{"en":{"e":"Safety + Longevity"},"jp":{"e":"安全性 + 持続性"},"cn":{"e":"安全性 + 持久性"},"tw":{"e":"安全性 + 持久性"},"th":{"e":"ความปลอดภัย + ความคงทน"}}},{"k":"부드러운 히알루론산 성분으로 민감한 부위에도 안정적으로 적용되어 자연스러운 결과와 지속력을 제공합니다","t":{"en":{"e":"With soft hyaluronic acid ingredients, it can be safely applied even to sensitive areas, providing natural results and long-lasting effects."},"jp":{"e":"高い安定性を持つフィラーを使用し、日常生活に大きな影響をほとんど与えず、自然なボリュームと弾力の形成をサポートします。"},"cn":{"e":"使用高稳定性的填充材料,在不影响日常生活的情况下,帮助打造自然的容积与弹性。"},"tw":{"e":"使用高穩定性的填充材料,在不影響日常生活的情況下,協助打造自然的容積與彈性。"},"th":{"e":"ด้วยสัมผัสที่นุ่มนวลและส่วนผสมที่เข้ากับร่างกาย จึงสามารถใช้กับบริเวณที่บอบบางได้อย่างปลอดภัย และให้ผลลัพธ์ที่ดูเป็นธรรมชาติ"}}},{"k":"약 20~30분 내외","t":{"en":{"e":"About 20-30 minutes"},"jp":{"e":"約20~30分"},"cn":{"e":"约20~30分钟"},"tw":{"e":"約20~30分鐘"},"th":{"e":"ประมาณ 20-30 นาที"}}},{"k":"시술 후 빠르게 회복","t":{"en":{"e":"Quick recovery after treatment"},"jp":{"e":"施術後、比較的早く回復"},"cn":{"e":"施术后可快速恢复"},"tw":{"e":"療程後可快速恢復"},"th":{"e":"กลับไปใช้ชีวิตประจำวันได้อย่างรวดเร็วหลังทำ"}}},{"k":"통증 정도","t":{"en":{"e":"Pain level"},"jp":{"e":"痛みの程度"},"cn":{"e":"疼痛程度"},"tw":{"e":"疼痛程度"},"th":{"e":"ความเจ็บ"}}},{"k":"비교적 적은 편","t":{"en":{"e":"Relatively mild"},"jp":{"e":"比較的少なめ"},"cn":{"e":"相对较轻"},"tw":{"e":"相對較低"},"th":{"e":"ค่อนข้างต่ำ"}}},{"k":"미미썸의원 시술 특징","t":{"en":{"e":"Key Features of MIMISOME Clinic Treatments"},"jp":{"e":"ミミソムクリニックの施術特徴"},"cn":{"e":"MIMISOME 诊所治疗特点"},"tw":{"e":"MIMISOME 診所療程特色"},"th":{"e":"จุดเด่นของหัตถการ MIMISOME"}}},{"k":"1:1 맞춤상담","t":{"en":{"e":"1:1 personalized consultation"},"jp":{"e":"1:1 オーダーメイド施術"},"cn":{"e":"1:1 定制施术"},"tw":{"e":"1:1 客製化療程"},"th":{"e":"ปรึกษาแบบตัวต่อตัว"}}},{"k":"노하우를 갖춘 의료진","t":{"en":{"e":"Experienced Medical Professionals"},"jp":{"e":"豊富な経験と高い技術力"},"cn":{"e":"丰富经验与技术力"},"tw":{"e":"豐富經驗與技術實力"},"th":{"e":"ทีมแพทย์ผู้มีประสบการณ์"}}},{"k":"심미성 및 기능성 향상","t":{"en":{"e":"Enhanced Aesthetics and Functionality"},"jp":{"e":"機能性と美しさを同時に向上"},"cn":{"e":"功能与美感同步提升"},"tw":{"e":"功能與美感同步提升"},"th":{"e":"เพิ่มทั้งความสวยงามและการใช้งาน"}}},{"k":"모공지우개 엑소좀 2.5cc","t":{"en":{"e":"Pore Care Exosome 2.5cc"},"jp":{"e":"毛穴ケア エクソソーム 2.5cc"},"tw":{"e":"毛孔護理 外泌體 2.5cc"}},"c":"스킨부스터"},{"k":"2.5cc 99,000원","t":{"en":{"e":"2.5cc 99,000 KRW"},"jp":{"e":"2.5cc 99,000 KRW"},"tw":{"e":"2.5cc 99,000 KRW"}}},{"k":"9.9만","t":{"en":{"e":"99,000 KRW"},"jp":{"e":"99,000 KRW"},"tw":{"e":"99,000 KRW"}}},{"k":"MIMISOME EXOSOME","t":{"en":{"e":"MIMISOME EXOSOME"},"jp":{"e":"MIMISOME EXOSOME"},"tw":{"e":"MIMISOME EXOSOME"}}},{"k":"근본적인 피부 개선 스킨부스터 엑소좀","t":{"en":{"e":"Fundamental Skin Improvement Skin Booster Exosome"},"jp":{"e":"根本からの肌改善 スキンブースター エクソソーム"},"tw":{"e":"從根本改善膚況 外泌體肌膚導入療程"}}},{"k":"콜라겐 재생","t":{"en":{"e":"Collagen Care"},"jp":{"e":"コラーゲンケア"},"tw":{"e":"膠原蛋白護理"}}},{"k":"유수분 밸런스","t":{"en":{"e":"Hydration Balance"},"jp":{"e":"水分バランス"},"tw":{"e":"水分平衡"}}},{"k":"모공 축소","t":{"en":{"e":"Pore Care"},"jp":{"e":"毛穴ケア"},"tw":{"e":"毛孔護理"}}},{"k":"미미썸 엑소좀 추천 대상","t":{"en":{"e":"MIMISOME Exosome Recommended For"},"jp":{"e":"ミミソム エクソソーム おすすめの方"},"tw":{"e":"MIMISOME 外泌體 適合對象"}}},{"k":"거칠고 푸석한 피부결이 고민이신 분","t":{"en":{"e":"Those concerned about rough, dry skin"},"jp":{"e":"肌のざらつき・乾燥が気になる方"},"tw":{"e":"在意肌膚粗糙、乾燥者"}}},{"k":"각종 트러블과 넓은 모공이 고민이신 분","t":{"en":{"e":"Those concerned about acne and enlarged pores"},"jp":{"e":"ニキビや毛穴の開きが気になる方"},"tw":{"e":"在意痘痘與毛孔粗大者"}}},{"k":"복합적인 피부 개선을 원하시는 분","t":{"en":{"e":"Those who want to improve multiple skin concerns"},"jp":{"e":"複合的な肌悩みを改善したい方"},"tw":{"e":"想改善多重肌膚問題者"}}},{"k":"무너진 피부 밸런스 케어","t":{"en":{"e":"Restore Your Skin Balance"},"jp":{"e":"崩れた肌バランスケア"},"tw":{"e":"重整失衡的肌膚狀態"}}},{"k":"미미썸만의 입체적인 모공 토탈 케어","t":{"en":{"e":"MIMISOME's Signature Multi-Dimensional Pore Care"},"jp":{"e":"ミミソムだけの 立体的な毛穴トータルケア"},"tw":{"e":"MIMISOME 專屬 立體毛孔全面護理"}}},{"k":"미미썸은 피부 상태를 정밀하게 분석하고 모공의 원인에 맞춰 시술을 설계합니다","t":{"en":{"e":"We carefully analyze your skin condition and design treatments based on the root causes of pores"},"jp":{"e":"ミミソムは肌状態を丁寧に分析し 毛穴の原因に合わせて施術を設計します"},"tw":{"e":"透過細緻的膚況分析 依毛孔成因規劃客製療程"}}},{"k":"단순 축소가 아닌 피부결, 탄력, 밀도까지 함께 개선합니다","t":{"en":{"e":"Not just basic skincare, but a total approach to texture, elasticity, and glow"},"jp":{"e":"肌キメ・ハリ・ツヤまでトータルに整えます"},"tw":{"e":"不只是基礎保養 從膚質、彈力到光澤全面調理"}}},{"k":"ASCE+ 핵심 성분 그대로","t":{"en":{"e":"ASCE+ Key Ingredients Intact"},"jp":{"e":"ASCE+ 主要成分そのまま"},"tw":{"e":"ASCE+ 核心成分完整保留"}}},{"k":"본 이벤트는 침습적 방식으로 시술하지 않습니다","t":{"en":{"e":"This event does not involve invasive procedures"},"jp":{"e":"本イベントは侵襲的な方法では施術を行いません"},"tw":{"e":"本活動非侵入式療程"}}},{"k":"ASCE+ 엑소좀 100%","t":{"en":{"e":"ASCE+ Exosome 100%"},"jp":{"e":"ASCE+ エクソソーム 100%"},"tw":{"e":"ASCE+ 外泌體 100%"}}},{"k":"콜라겐 600%","t":{"en":{"e":"Collagen 600%"},"jp":{"e":"コラーゲン 600%"},"tw":{"e":"膠原蛋白 600%"}}},{"k":"엘라스틴 300%","t":{"en":{"e":"Elastin 300%"},"jp":{"e":"エラスチン 300%"},"tw":{"e":"彈力蛋白 300%"}}},{"k":"MIMISOME 오시는 길","t":{"en":{"e":"Directions to MIMISOME"},"jp":{"e":"MIMISOME へのアクセス"},"tw":{"e":"MIMISOME 交通指南"}}},{"k":"진료시간안내","t":{"en":{"e":"Operating Hours"},"jp":{"e":"診療時間のご案内"},"tw":{"e":"看診時間"}}},{"k":"평일 10:00 - 20:00","t":{"en":{"e":"Weekdays 10:00 - 20:00"},"jp":{"e":"平日 10:00 - 20:00"},"tw":{"e":"平日 10:00 - 20:00"}}},{"k":"토요일 10:00 - 16:00","t":{"en":{"e":"Saturday 10:00 - 16:00"},"jp":{"e":"土曜日 10:00 - 16:00"},"tw":{"e":"週六 10:00 - 16:00"}}},{"k":"점심시간 13:00 - 14:00","t":{"en":{"e":"Lunch break 13:00 - 14:00"},"jp":{"e":"ランチタイム 13:00 - 14:00"},"tw":{"e":"午休時間 13:00 - 14:00"}}},{"k":"공휴일 / 일요일 휴진","t":{"en":{"e":"Closed on holidays and Sundays"},"jp":{"e":"祝日・日曜日休診"},"tw":{"e":"國定假日／週日休診"}}},{"k":"팅커벨 귀필러","t":{"en":{"e":"Tinker Bell Ear Filler"},"jp":{"e":"ティンカーベル耳フィラー"},"cn":{"e":"小精灵耳部填充"},"tw":{"e":"精靈耳玻尿酸"},"th":{"e":"ฉีดฟิลเลอร์หู"}},"c":"필러"},{"k":"팅커벨귀필러","t":{"en":{"e":"Tinker Bell Ear Filler"},"jp":{"e":"ティンカーベル耳フィラー"},"cn":{"e":"小精灵耳部填充"},"tw":{"e":"精靈耳玻尿酸"},"th":{"e":"ฉีดฟิลเลอร์หู"}},"c":"필러"},{"k":"165,000원","t":{"en":{"e":"165,000 KRW"},"jp":{"e":"165,000 KRW"},"cn":{"e":"165,000 KRW"},"tw":{"e":"165,000 KRW"},"th":{"e":"165,000 KRW"}}},{"k":"작아보이는 얼굴의 비결","t":{"en":{"e":"The secret to a smaller looking face"},"jp":{"e":"小顔に見せる秘訣"},"cn":{"e":"让脸看起来更小的秘诀"},"tw":{"e":"讓臉看起來更小的秘密"},"th":{"e":"เคล็ดลับ ให้หน้าดูเล็กเวลาถ่ายรูป"}}},{"k":"비대칭 귓볼 & 얇은 귓볼","t":{"en":{"e":"Asymmetrical earlobes and thin earlobes"},"jp":{"e":"非対称の耳たぶと薄い耳たぶ"},"cn":{"e":"不对称的耳垂与薄耳垂"},"tw":{"e":"不對稱耳垂與薄耳垂"},"th":{"e":"ติ่งหูไม่สมมาตร และติ่งหูบาง"}}},{"k":"칼 귀","t":{"en":{"e":"Sharp ears"},"jp":{"e":"尖った耳"},"cn":{"e":"尖耳"},"tw":{"e":"尖耳"},"th":{"e":"หูแหลม"}}},{"k":"울퉁불퉁한 귀","t":{"en":{"e":"Uneven ears"},"jp":{"e":"不均一な耳"},"cn":{"e":"凹凸不平的耳朵"},"tw":{"e":"不平整的耳朵"},"th":{"e":"หูไม่เรียบ"}}},{"k":"누운 귀","t":{"en":{"e":"Drooping ears"},"jp":{"e":"垂れた耳"},"cn":{"e":"贴头型耳朵"},"tw":{"e":"貼頭型耳朵"},"th":{"e":"หูแนบศีรษะ"}}},{"k":"작은 귀","t":{"en":{"e":"Small ears"},"jp":{"e":"小さい耳"},"cn":{"e":"小耳朵"},"tw":{"e":"小耳朵"},"th":{"e":"หูเล็ก"}}},{"k":"미미썸 팅커벨 귀필러를 통해 개선할 수 있습니다.","t":{"en":{"e":"Can be improved with MIMISOME Tinker Bell Ear Filler"},"jp":{"e":"ミミソムのティンカーベル耳フィラーで改善できます"},"cn":{"e":"可通过 Mimisome 小精灵耳部填充疗程改善"},"tw":{"e":"精靈耳Mimisome 小精靈耳朵填充療程可改善"},"th":{"e":"สามารถปรับปรุงได้ด้วยฟิลเลอร์หูทิงเกอร์เบลล์ของ Mimisome"}}},{"k":"가장 이상적인 귀 비율","t":{"en":{"e":"Most ideal ear proportions"},"jp":{"e":"最も理想的な耳の比率"},"cn":{"e":"最理想的耳部比例"},"tw":{"e":"最理想的耳朵比例"},"th":{"e":"สัดส่วนหูที่เหมาะสมที่สุด"}}},{"k":"예쁜 귀 길이","t":{"en":{"e":"Pretty length"},"jp":{"e":"美しい長さ"},"cn":{"e":"理想耳长"},"tw":{"e":"理想耳長"},"th":{"e":"ความยาวหูที่สวยงาม"}}},{"k":"여성 5.5cm-6cm | 남성 7cm","t":{"en":{"e":"Women 5.5 cm to 6 cm, Men 7 cm"},"jp":{"e":"女性5.5cmから6cm, 男性7cm"},"cn":{"e":"女性 5.5-6 厘米 | 男性 7 厘米"},"tw":{"e":"女性 5.5–6 公分 | 男性 7 公分"},"th":{"e":"ผู้หญิง 5.5-6 ซม. | ผู้ชาย 7 ซม."}}},{"k":"귓볼 크기","t":{"en":{"e":"Earlobe size"},"jp":{"e":"耳たぶの大きさ"},"cn":{"e":"耳垂大小"},"tw":{"e":"耳垂大小"},"th":{"e":"ขนาดติ่งหู"}}},{"k":"귀 전체 크기의 25%","t":{"en":{"e":"25 percent of the entire ear"},"jp":{"e":"耳全体の25パーセント"},"cn":{"e":"约占整体耳朵的 25%"},"tw":{"e":"約佔整體耳朵的 25%"},"th":{"e":"25% ของขนาดหูทั้งหมด"}}},{"k":"귀가 작고 누워있어 정면에서 보이지 않고 얼굴 비율이 마음에 들지 않는 경우","t":{"en":{"e":"If your ears are small and laid back, not visible from the front, and your facial proportions are not satisfying"},"jp":{"e":"耳が小さく後ろに寝ていて正面から見えにくい場合 顔のバランスに満足できない場合"},"cn":{"e":"若耳朵较小且贴头, 从正面不明显, 导致脸部比例不理想"},"tw":{"e":"若耳朵較小且貼頭, 正面看不清楚, 臉部比例不理想時"},"th":{"e":"หากหูเล็ก แนบศีรษะ มองจากด้านหน้าไม่ค่อยเห็น และสัดส่วนใบหน้าไม่ถูกใจ"}}},{"k":"귀 모양 개선만으로 얼굴 길이와 크기 축소 효과 이목구비의 균형 및 얼굴을 조화롭게 만들어주는 시술입니다.","t":{"en":{"e":"This procedure improves ear shape to make the face look smaller and create more balanced facial proportions"},"jp":{"e":"耳の形を整え 顔をより小さく見せ バランスの取れたフェイスラインへ導きます"},"cn":{"e":"仅改善耳型 即可达到缩短脸部长度与比例的视觉效果 让五官更加平衡, 使整体脸型更协调"},"tw":{"e":"僅改善耳型 即可達到縮短臉部長度與比例的視覺效果, 讓五官更加平衡, 使整體臉型更協調"},"th":{"e":"เพียงปรับรูปทรงหู สามารถช่วยให้ใบหน้าดูสั้นลงและเลดลง สร้างสมดุลของโครงหน้า และทำให้ใบหน้าดูสวยกลมกลืนมากขึ้น"}}},{"k":"소얼굴 효과 발군","t":{"en":{"e":"Outstanding face slimming effect"},"jp":{"e":"小顔効果抜群"},"cn":{"e":"小脸效果出众"},"tw":{"e":"小顏效果拔群"},"th":{"e":"เห็นผลหน้าเล็กชัดเจน"}}},{"k":"화제의 팅커벨, 순식간에 예뻐지게 해드려요","t":{"en":{"e":"Trending Tinker Bell ears make you beautiful in an instant"},"jp":{"e":"話題のティンカーベル耳で一瞬で美しく"},"cn":{"e":"话题精灵耳 让你瞬间变美"},"tw":{"e":"話題精靈耳 讓你瞬間變美"},"th":{"e":"ทิงเกอร์เบลล์สุดฮอต ทำให้คุณสวยในทันที"}}},{"k":"써마지 FLX 300샷","c":"리프팅","t":{"en":{"e":"Thermage FLX 300 Shots"},"jp":{"e":"サーマクールFLX 300ショット"},"tw":{"e":"第四代鳳凰電波 300發"}}},{"k":"써마지 FLX","c":"리프팅","t":{"en":{"e":"Thermage FLX"},"jp":{"e":"サーマクールFLX"},"tw":{"e":"第四代鳳凰電波"}}},{"k":"써마지FLX","c":"리프팅","t":{"en":{"e":"Thermage FLX"},"jp":{"e":"サーマクールFLX"},"tw":{"e":"第四代鳳凰電波"}}},{"k":"피부 속부터 탱탱하게 차오르는","t":{"en":{"e":"Firming up from deep within the skin"},"jp":{"e":"肌の内側からふっくらと満ちる"},"tw":{"e":"由肌底深層開始飽滿緊緻"}}},{"k":"고주파 리프팅","t":{"en":{"e":"RF Lifting"},"jp":{"e":"高周波リフティング"},"tw":{"e":"電波拉提"}}},{"k":"고주파 에너지로 피부 탄력 개선","t":{"en":{"e":"Improves skin elasticity with radiofrequency energy"},"jp":{"e":"高周波エネルギーで肌のハリを改善"},"tw":{"e":"透過電波能量改善肌膚彈性"}}},{"k":"피부 상태에 맞춘 맞춤 리프팅","t":{"en":{"e":"Customized lifting tailored to your skin condition"},"jp":{"e":"肌状態に合わせたオーダーメイドリフティング"},"tw":{"e":"依照肌膚狀況打造客製化拉提"}}},{"k":"처진 윤곽과 잔주름까지 탄탄하게","t":{"en":{"e":"Helps firm sagging contours and smooth fine lines"},"jp":{"e":"たるんだフェイスラインや小ジワまでしっかりケア"},"tw":{"e":"改善鬆弛輪廓與細紋,使肌膚更緊實"}}},{"k":"전방향 진동으로 통증 부담 완화","t":{"en":{"e":"Reduces discomfort with multidirectional vibration technology"},"jp":{"e":"全方向振動で痛みの負担を軽減"},"tw":{"e":"以全方位震動技術降低疼痛負擔"}}},{"k":"무너진 탄력이 고민이라면?","t":{"en":{"e":"Concerned about loss of firmness?"},"jp":{"e":"失われたハリが気になるなら?"},"tw":{"e":"如果您正煩惱肌膚彈力流失?"}}},{"k":"탄력개선","t":{"en":{"e":"Firming"},"jp":{"e":"ハリ改善"},"tw":{"e":"彈力改善"}}},{"k":"윤곽정리","t":{"en":{"e":"Contour definition"},"jp":{"e":"輪郭補整"},"tw":{"e":"輪廓緊緻"}}},{"k":"잔주름 완화","t":{"en":{"e":"Fine line reduction"},"jp":{"e":"小ジワ改善"},"tw":{"e":"淡化細紋"}}},{"k":"스킨 타이트닝","t":{"en":{"e":"Skin tightening"},"jp":{"e":"スキンタイトニング"},"tw":{"e":"肌膚緊實"}}},{"k":"이런 분들께 추천드려요","t":{"en":{"e":"Recommended for those who"},"jp":{"e":"このような方におすすめです"},"tw":{"e":"推薦給這樣的您"}}},{"k":"얼굴 전체 탄력과 잔주름 개선을 원하시는 분","t":{"en":{"e":"Want to improve overall facial firmness and fine lines"},"jp":{"e":"顔全体のハリや小ジワを改善したい方"},"tw":{"e":"想改善全臉彈力與細紋的人"}}},{"k":"리프팅 시술 시 통증이 부담스러우신 분","t":{"en":{"e":"Are concerned about discomfort during lifting treatments"},"jp":{"e":"リフティング施術の痛みが気になる方"},"tw":{"e":"對拉提療程疼痛感到有負擔的人"}}},{"k":"늘어진 턱선과 얼굴라인이 고민이신 분","t":{"en":{"e":"Are concerned about a sagging jawline and facial contours"},"jp":{"e":"たるんだフェイスラインやあご周りが気になる方"},"tw":{"e":"在意下顎線鬆弛與臉部輪廓下垂的人"}}},{"k":"입가·눈가 주름이 심해 나이 들어보이시는 분","t":{"en":{"e":"Have noticeable wrinkles around the mouth and eyes that make them look older"},"jp":{"e":"口元・目元のシワが目立ち、老けて見えやすい方"},"tw":{"e":"因嘴角、眼周皺紋明顯而顯老的人"}}},{"k":"#탄력UP","t":{"en":{"e":"#Firmness UP"},"jp":{"e":"#ハリUP"},"tw":{"e":"#彈力提升"}}},{"k":"#또렷한 윤곽","t":{"en":{"e":"#Defined contours"},"jp":{"e":"#くっきり輪郭"},"tw":{"e":"#輪廓更明顯"}}},{"k":"#자연스러운 리프팅","t":{"en":{"e":"#Natural-looking lifting"},"jp":{"e":"#自然なリフティング"},"tw":{"e":"#自然拉提效果"}}},{"k":"해당 사진은 미미썸의원의 광고용 이미지 모델입니다","t":{"en":{"e":"This image is for advertising purposes only for MIMISOME Clinic"},"jp":{"e":"本画像はミミサム医院の広告用イメージです"},"tw":{"e":"本圖片僅作為米米森醫診所廣告形象示意使用"}}},{"k":"1,210,000원","t":{"en":{"e":"1,210,000 KRW"},"jp":{"e":"1,210,000 KRW"},"tw":{"e":"1,210,000 KRW"}}},{"k":"VAT INCLUDED","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"tw":{"e":"含VAT"}}},{"k":"*VAT INCLUDED","t":{"en":{"e":"*VAT INCLUDED"},"jp":{"e":"*VAT込み"},"tw":{"e":"*含VAT"}}},{"k":"*VAT込み","t":{"en":{"e":"*VAT INCLUDED"},"jp":{"e":"*VAT込み"},"tw":{"e":"*含VAT"}}},{"k":"*含VAT","t":{"en":{"e":"*VAT INCLUDED"},"jp":{"e":"*VAT込み"},"tw":{"e":"*含VAT"}}},{"k":"포텐자 레이저","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza Laser"},"jp":{"e":"ポテンツァレーザー"}}},{"k":"포텐자레이저","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza Laser"},"jp":{"e":"ポテンツァレーザー"}}},{"k":"미미썸 포텐자","t":{"en":{"e":"MIMISOME Potenza"},"jp":{"e":"ミミサム ポテンツァ"}}},{"k":"미미썸 포텐자 레이저","t":{"en":{"e":"MIMISOME Potenza Laser"},"jp":{"e":"ミミサム ポテンツァレーザー"}}},{"k":"#모공축소","t":{"en":{"e":"#PoreReduction"},"jp":{"e":"#毛穴縮小"}}},{"k":"#흉터치료","t":{"en":{"e":"#ScarTreatment"},"jp":{"e":"#傷跡治療"}}},{"k":"#탄력개선","t":{"en":{"e":"#ElasticityImprovement"},"jp":{"e":"#ハリ改善"}}},{"k":"얼굴에 여드름 흉터가 남았어","t":{"en":{"e":"I have acne scars left on my face"},"jp":{"e":"顔にニキビ跡が残っちゃった"}}},{"k":"나이가 들수록 모공이 안줄어드네","t":{"en":{"e":"My pores don't seem to shrink"},"jp":{"e":"年を重ねるごとに毛穴が目立つ"}}},{"k":"피부에 탄력 좀 주고 싶은데","t":{"en":{"e":"I want to improve my skin elasticity"},"jp":{"e":"肌にハリを出したいのに"}}},{"k":"어느새 생긴 모공&여드름흉터가 고민이라면","t":{"en":{"e":"If you're concerned about enlarged pores and acne scars that appeared over time"},"jp":{"e":"いつの間にかできた毛穴やニキビ跡が気になるなら"}}},{"k":"미미썸 포텐자로 해결해보세요!","t":{"en":{"e":"try solving it with Potenza!"},"jp":{"e":"ミミサム ポテンツァで解決してみてください!"}}},{"k":"미미썸 포텐자로 모공&여드름흉터 해결해요!","t":{"en":{"e":"Pores & acne scars? Potenza."},"jp":{"e":"ミミサム・ポテンツァで毛穴&ニキビ跡解決してみて!"}}},{"k":"미미썸 포텐자레이저는 고주파 종류와 깊이를 선택해 다양한 피부층에 에너지를 전달합니다","t":{"en":{"e":"Delivers energy to various skin layers by adjusting RF type and depth"},"jp":{"e":"ミミサム・ポテンツァレーザーは高周波の種類と深さを選択し、さまざまな皮膚層へエネルギーを届けます"}}},{"k":"시술횟수는 개개인의 피부 상태에 따라 상이하며, 적절한 시술횟수는 상담 후 확인하실 수 있습니다","t":{"en":{"e":"Sessions vary by skin condition. Final plan decided after consultation"},"jp":{"e":"施術回数は個人の肌状態によって異なり、適切な施術回数はカウンセリング後にご案内いたします"}}},{"k":"#깐달걀피부","t":{"en":{"e":"#GlassSkin"},"jp":{"e":"#ゆで卵肌"}}},{"k":"#자신감UP","t":{"en":{"e":"#ConfidenceUP"},"jp":{"e":"#自信アップ"}}},{"k":"#미미썸과상담","t":{"en":{"e":"#ConsultMIMISOME"},"jp":{"e":"#ミミサム相談"}}},{"k":"포텐자가 필요해","t":{"en":{"e":"You need Potenza"},"jp":{"e":"ポテンツァが必要かも"}}},{"k":"이런 분들에게 추천해드려요!","t":{"en":{"e":"Recommended for"},"jp":{"e":"こんな方におすすめです!"}}},{"k":"피부 탄력이 떨어져 잔주름이 고민이신 분","t":{"en":{"e":"Concerned about wrinkles from loss of elasticity"},"jp":{"e":"肌のハリが低下し・小ジワが気になる方"}}},{"k":"자연스러운 리프팅, 타이트닝을 원하시는 분","t":{"en":{"e":"Those who want natural lifting and tightening"},"jp":{"e":"自然なリフトアップ・タイトニングを求める方"}}},{"k":"얼굴 라인, 모공, 피부결이 고민이신 분","t":{"en":{"e":"Concerned about contour, pores, and texture"},"jp":{"e":"フェイスライン・毛穴・キメが気になる方"}}},{"k":"예민해진 피부를 재생시키고 싶으신 분","t":{"en":{"e":"Those who want to restore sensitive or damaged skin"},"jp":{"e":"敏感になった肌を再生したい方"}}},{"k":"*카운셀링 포함","t":{"en":{"e":"*CONSULTATION INCLUDED"},"jp":{"e":"※カウンセリング込み"}}},{"k":"※カウンセリング込み","t":{"en":{"e":"*CONSULTATION INCLUDED"},"jp":{"e":"※カウンセリング込み"}}},{"k":"※VAT込み","t":{"en":{"e":"*VAT INCLUDED"},"jp":{"e":"※VAT込み"}}},{"k":"209,000원","t":{"en":{"e":"209,000 KRW"},"jp":{"e":"209,000 KRW"}}},{"k":"리쥬란 HB PLUS","t":{"en":{"e":"REJURAN HB PLUS"},"jp":{"e":"リジュラン HB PLUS"},"cn":{"e":"丽珠兰 HB PLUS"},"tw":{"e":"麗珠蘭 HB PLUS"},"th":{"e":"REJURAN HB PLUS"}},"c":"스킨부스터"},{"k":"리쥬란 HB plus","t":{"en":{"e":"REJURAN HB plus"},"jp":{"e":"リジュラン HB plus"},"cn":{"e":"丽珠兰 HB plus"},"tw":{"e":"麗珠蘭 HB plus"},"th":{"e":"รีจูรัน HB plus"}},"c":"스킨부스터"},{"k":"리쥬란HB plus","t":{"en":{"e":"REJURAN HB plus"},"jp":{"e":"リジュランHB plus"},"cn":{"e":"丽珠兰 HB plus"},"tw":{"e":"麗珠蘭 HB plus"},"th":{"e":"รีจูรัน HB plus"}},"c":"스킨부스터"},{"k":"리쥬란HB+","t":{"en":{"e":"Rejuran HB+"},"jp":{"e":"リジュランHB+"},"cn":{"e":"丽珠兰 HB+"},"tw":{"e":"麗珠蘭 HB+"},"th":{"e":"รีจูรัน HB+"}},"c":"스킨부스터"},{"k":"피부 속부터 탄탄하게","t":{"en":{"e":"Firming from deep within the skin"},"jp":{"e":"肌の内側からしっかりと整える"},"cn":{"e":"从肌肤深层开始紧致"},"tw":{"e":"從肌膚深層開始緊緻"},"th":{"e":"ฟื้นฟูผิวให้แข็งแรงจากชั้นใน"}}},{"k":"가장 근본적인 안티에이징!","t":{"en":{"e":"The most fundamental anti-aging solution"},"jp":{"e":"根本的なアンチエイジング"},"cn":{"e":"根本性抗衰老"},"tw":{"e":"根本性抗老化"},"th":{"e":"การชะลอวัยจากต้นตอ"}}},{"k":"피부 깊은 곳에서부터 차오르는 스킨부스터","t":{"en":{"e":"A skin booster that revitalizes the skin from deep within"},"jp":{"e":"肌の奥深くから満ちあふれるスキンブースター"},"cn":{"e":"从肌肤深层焕发的肌肤增强针"},"tw":{"e":"從肌膚深層湧現的肌膚強化針"},"th":{"e":"สกินบูสเตอร์ที่ฟื้นฟูจากชั้นผิวลึก"}}},{"k":"#피부재생","t":{"en":{"e":"#Skin regeneration"},"jp":{"e":"#皮膚再生"},"cn":{"e":"#皮肤再生"},"tw":{"e":"#皮膚再生"},"th":{"e":"#ฟื้นฟูผิว"}}},{"k":"#탄력상승","t":{"en":{"e":"#Improved elasticity"},"jp":{"e":"#弾力アップ"},"cn":{"e":"#弹力提升"},"tw":{"e":"#彈力提升"},"th":{"e":"#เพิ่มความกระชับ"}}},{"k":"#피부결 개선","t":{"en":{"e":"#Improved skin texture"},"jp":{"e":"#肌キメ改善"},"cn":{"e":"#肤质改善"},"tw":{"e":"#膚質改善"},"th":{"e":"#ผิวเรียบเนียนขึ้น"}}},{"k":"얇아진 피부, 피부 속부터 개선","t":{"en":{"e":"Improves thinned skin from within"},"jp":{"e":"薄くなった肌 肌の内側から改善"},"cn":{"e":"变薄的肌肤，从肌肤内部开始改善"},"tw":{"e":"變薄的肌膚，從肌膚內部開始改善"},"th":{"e":"ผิวที่บางลง ฟื้นฟูจากภายในผิว"}}},{"k":"약해진 피부 속 깊은 곳부터 차오르는 재생","t":{"en":{"e":"Promotes regeneration deep within weakened skin"},"jp":{"e":"弱くなった肌の奥深くから 湧き上がる再生"},"cn":{"e":"脆弱的肌肤内部，从深层涌现的新生"},"tw":{"e":"脆弱的肌膚內部，從深層湧現的新生"},"th":{"e":"ผิวที่อ่อนแอ ฟื้นฟูจากชั้นลึกภายใน"}}},{"k":"깨끗하고 맑은 피부로 돌아가요!","t":{"en":{"e":"Helps restore clear, healthy-looking skin!"},"jp":{"e":"クリアで透明感のある肌へ戻りましょう！"},"cn":{"e":"重回洁净透亮的肌肤！"},"tw":{"e":"重回潔淨透亮的肌膚！"},"th":{"e":"กลับสู่ผิวใสสะอาดอีกครั้ง!"}}},{"k":"리쥬란이란?","t":{"en":{"e":"What is REJURAN?"},"jp":{"e":"リジュランHB+とは？"},"cn":{"e":"什么是丽珠兰？"},"tw":{"e":"什麼是麗珠蘭？"},"th":{"e":"รีจูรันคืออะไร?"}}},{"k":"연어주사라고 불리는 리쥬란은 연어에서 추출한 폴리뉴클레오타이드(PN)성분을 피부에 주입해 피부 속 환경을 개선해 주는 스킨부스터 시술입니다.","t":{"en":{"e":"Rejuran, also known as the \\"salmon injection,\\" is a skin booster treatment that uses polynucleotide (PN) extracted from salmon to improve the skin's internal environment."},"jp":{"e":"サーモン由来のポリヌクレオチド(PN)成分を肌に注入し、肌の内側の環境を改善するスキンブースターです。"},"cn":{"e":"将从鲑鱼中提取的多核苷酸 (PN) 成分注入肌肤，改善肌肤内部环境的肌肤增强针。"},"tw":{"e":"將從鮭魚中萃取的多核苷酸 (PN) 成分注入肌膚，改善肌膚內部環境的肌膚強化針。"},"th":{"e":"สกินบูสเตอร์ที่ฉีดสารโพลีนิวคลีโอไทด์ (PN) ซึ่งสกัดจากปลาแซลมอนเข้าสู่ผิว เพื่อฟื้นฟูสภาพผิวจากภายใน"}}},{"k":"인체에 안전한 생체 적합 물질 PN 성분","t":{"en":{"e":"Biocompatible PN ingredients that are safe for the human body"},"jp":{"e":"人体に安全な生体適合PN成分"},"cn":{"e":"对人体安全的生物相容PN成分"},"tw":{"e":"對人體安全的生物相容PN成分"},"th":{"e":"สาร PN ที่เข้ากันได้กับร่างกายและปลอดภัย"}}},{"k":"일정한 형태를 유지하며 서서히 분해","t":{"en":{"e":"Maintains a consistent structure and gradually breaks down over time"},"jp":{"e":"一定の形状を保ちながらゆっくり分解"},"cn":{"e":"保持固定形态并逐渐分解"},"tw":{"e":"保持固定形態並逐漸分解"},"th":{"e":"คงรูปทรงและสลายตัวอย่างช้าๆ"}}},{"k":"통증은 Down","t":{"en":{"e":"LESS PAIN Down"},"jp":{"e":"痛みは Down"},"cn":{"e":"疼痛是 Down"},"tw":{"e":"疼痛是 Down"},"th":{"e":"ความเจ็บปวดคือ Down"}}},{"k":"효과는 UP","t":{"en":{"e":"ENHANCED RESULTS UP"},"jp":{"e":"効果は UP"},"cn":{"e":"效果是 UP"},"tw":{"e":"效果是 UP"},"th":{"e":"ผลลัพธ์คือ UP"}}},{"k":"기존 리쥬란 힐러의 히알루론산과 국소마취성분(리도카인)이 더해져 통증은 줄이고 수분감은 업그레이드 시켰습니다.","t":{"en":{"e":"With the addition of hyaluronic acid and the local anesthetic lidocaine to the original Rejuran Healer formula, discomfort is reduced while hydration is enhanced."},"jp":{"e":"リジュランシリーズにヒアルロン酸と局所麻酔成分リドカインを配合し、痛みを軽減し保湿力をアップグレードしました。"},"cn":{"e":"在原有丽珠兰系列基础上添加透明质酸，并加入局部麻醉成分利多卡因，降低疼痛感并升级保湿效果。"},"tw":{"e":"在原有麗珠蘭系列基礎上添加玻尿酸，並加入局部麻醉成分利多卡因，降低疼痛感並升級保濕效果。"},"th":{"e":"เพิ่มกรดไฮยาลูรอนิกและสารยาชาเฉพาะที่ลิโดเคนในรีจูรันเดิม ช่วยลดความเจ็บและเพิ่มความชุ่มชื้นให้ดียิ่งขึ้น"}}},{"k":"히알루론산","t":{"en":{"e":"Hyaluronic Acid"},"jp":{"e":"ヒアルロン酸"},"cn":{"e":"透明质酸"},"tw":{"e":"玻尿酸"},"th":{"e":"กรดไฮยาลูรอนิก"}}},{"k":"리쥬란 힐러","t":{"en":{"e":"Rejuran Healer"},"jp":{"e":"リジュラン成分"},"cn":{"e":"丽珠兰成分"},"tw":{"e":"麗珠蘭成分"},"th":{"e":"สารรีจูรัน"}}},{"k":"리도카인","t":{"en":{"e":"Lidocaine"},"jp":{"e":"リドカイン"},"cn":{"e":"利多卡因"},"tw":{"e":"利多卡因"},"th":{"e":"ลิโดเคน"}}},{"k":"이런 분들께 추천합니다!","t":{"en":{"e":"Recommended for those who!"},"jp":{"e":"こんな方におすすめです！"},"cn":{"e":"推荐给以下人群！"},"tw":{"e":"推薦給以下族群！"},"th":{"e":"เหมาะสำหรับผู้ที่มีปัญหาเหล่านี้!"}}},{"k":"피부가 건조하고 푸석해진 분","t":{"en":{"e":"Have dry, rough skin"},"jp":{"e":"肌が乾燥してカサついている方"},"cn":{"e":"皮肤干燥粗糙的人"},"tw":{"e":"肌膚乾燥粗糙的人"},"th":{"e":"ผู้ที่ผิวแห้งและหยาบกร้าน"}}},{"k":"탄력 저하와 잔주름이 생긴 분","t":{"en":{"e":"Are experiencing loss of elasticity and fine lines"},"jp":{"e":"弾力低下や小じわが気になる方"},"cn":{"e":"出现弹力下降和细纹的人"},"tw":{"e":"出現彈力下降與細紋的人"},"th":{"e":"ผู้ที่มีความกระชับลดลงและริ้วรอยเล็กๆ"}}},{"k":"윤기와 생기가 부족한 피부인 분","t":{"en":{"e":"Have dull, tired-looking skin lacking radiance and vitality"},"jp":{"e":"ツヤとハリが不足している肌の方"},"cn":{"e":"缺乏光泽与活力的肌肤"},"tw":{"e":"缺乏光澤與活力的肌膚"},"th":{"e":"ผู้ที่ผิวขาดความเปล่งปลั่งและความสดใส"}}},{"k":"모공이 넓고 피부결이 거친 분","t":{"en":{"e":"Have enlarged pores and uneven skin texture"},"jp":{"e":"毛穴が目立ち、肌キメが荒れている方"},"cn":{"e":"毛孔粗大、肤质粗糙的人"},"tw":{"e":"毛孔粗大、膚質粗糙的人"},"th":{"e":"ผู้ที่มีรูขุมขนกว้างและผิวไม่เรียบ"}}},{"k":"보습과 탄력을 동시에 원하는 분","t":{"en":{"e":"Want both hydration and improved elasticity"},"jp":{"e":"保湿と弾力を同時に求める方"},"cn":{"e":"希望同时改善保湿与弹力的人"},"tw":{"e":"希望同時改善保濕與彈力的人"},"th":{"e":"ผู้ที่ต้องการเพิ่มทั้งความชุ่มชื้นและความกระชับพร้อมกัน"}}},{"k":"99,000원","t":{"en":{"e":"99,000 KRW"},"jp":{"e":"99,000 KRW"},"cn":{"e":"99,000 KRW"},"tw":{"e":"99,000 KRW"},"th":{"e":"99,000 KRW"}}},{"k":"含增值税","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"含增值稅","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"รวม VAT แล้ว","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"쥬베룩","t":{"en":{"e":"JUVELOOK / Juvelook"},"jp":{"e":"ジュベルック"},"cn":{"e":"乔雅露"},"tw":{"e":"喬雅露"},"th":{"e":"Juvelook"}},"c":"스킨부스터"},{"k":"쥬베룩 3cc","t":{"en":{"e":"JUVELOOK 3cc"},"jp":{"e":"ジュベルック 3cc"},"cn":{"e":"乔雅露 3cc"},"tw":{"e":"喬雅露 3cc"},"th":{"e":"Juvelook 3cc"}},"c":"스킨부스터"},{"k":"쥬베룩 스킨 3cc","t":{"en":{"e":"JUVELOOK Skin 3cc"},"jp":{"e":"ジュベルックスキン 3cc"},"cn":{"e":"乔雅露 3cc"},"tw":{"e":"JUVELOOK喬雅露 3cc"},"th":{"e":"JUVELOOK 3cc"}},"c":"스킨부스터"},{"k":"피부를 건강하고 탄력있게 만들어주는","t":{"en":{"e":"Promotes healthier, firmer-looking skin"},"jp":{"e":"肌を健康でハリのある状態へ導く"},"cn":{"e":"打造健康紧致肌肤的"},"tw":{"e":"讓肌膚健康有彈力的"},"th":{"e":"ช่วยให้ผิวดูสุขภาพดีและกระชับ"}}},{"k":"자가콜라겐 생성 촉진","t":{"en":{"e":"Stimulates natural collagen production"},"jp":{"e":"自己コラーゲン生成促進"},"cn":{"e":"促进自身胶原蛋白生成"},"tw":{"e":"促進自身膠原蛋白生成"},"th":{"e":"กระตุ้นการสร้างคอลลาเจนตามธรรมชาติ"}}},{"k":"피부 탄력 및 잔주름 개선, 모공 개선","t":{"en":{"e":"Helps improve skin elasticity, fine lines, and the appearance of pores"},"jp":{"e":"肌のハリ向上・小じわ改善・毛穴改善"},"cn":{"e":"改善皮肤弹力与细纹，改善毛孔"},"tw":{"e":"改善肌膚彈力與細紋，改善毛孔"},"th":{"e":"ความกระชับของผิวและลดริ้วรอยเล็กๆ"}}},{"k":"탄력과 물광을 동시에 원하시나요?","t":{"en":{"e":"Are you looking for both firmness and a radiant glow?"},"jp":{"e":"ハリとツヤを同時に求めていますか？"},"cn":{"e":"想同时拥有紧致与水光肌吗？"},"tw":{"e":"想同時擁有彈力與水光肌嗎？"},"th":{"e":"ต้องการทั้งความกระชับและผิวฉ่ำโกลว์พร้อมกันหรือไม่?"}}},{"k":"피부재생","t":{"en":{"e":"Skin rejuvenation"},"jp":{"e":"肌再生"},"cn":{"e":"肌肤再生"},"tw":{"e":"肌膚再生"},"th":{"e":"ฟื้นฟูผิว"}}},{"k":"잔주름 개선","t":{"en":{"e":"Improvement of fine lines"},"jp":{"e":"小じわ改善"},"cn":{"e":"改善细纹"},"tw":{"e":"改善細紋"},"th":{"e":"ลดริ้วรอยเล็กๆ"}}},{"k":"자가콜라겐 생성","t":{"en":{"e":"Natural collagen stimulation"},"jp":{"e":"自己コラーゲン生成"},"cn":{"e":"促进胶原生成"},"tw":{"e":"促進膠原生成"},"th":{"e":"คอลลาเจนตามธรรมชาติถูกกระตุ้น"}}},{"k":"넓은 모공 개선","t":{"en":{"e":"Improvement of enlarged pores"},"jp":{"e":"開いた毛穴改善"},"cn":{"e":"改善粗大毛孔"},"tw":{"e":"改善粗大毛孔"},"th":{"e":"ช่วยกระชับรูขุมขนกว้าง"}}},{"k":"이런 분들에게 추천해요.","t":{"en":{"e":"Recommended for those who"},"jp":{"e":"こんな方におすすめです."},"cn":{"e":"推荐给以下人群."},"tw":{"e":"推薦給這些人."},"th":{"e":"แนะนำสำหรับผู้ที่"}}},{"k":"건강한 피부를 원하시는 분","t":{"en":{"e":"Want healthier-looking skin"},"jp":{"e":"健康的な肌を目指したい方"},"cn":{"e":"希望拥有健康肌肤的人"},"tw":{"e":"想要健康肌膚的人"},"th":{"e":"ผู้ที่ต้องการผิวสุขภาพดี"}}},{"k":"피부가 푸석하고 거칠며, 탄력이 떨어지신 분","t":{"en":{"e":"Have dull, rough skin with reduced elasticity"},"jp":{"e":"肌が乾燥してごわつき、ハリが低下している方"},"cn":{"e":"皮肤干燥粗糙、弹力下降的人"},"tw":{"e":"肌膚乾燥粗糙、彈力下降的人"},"th":{"e":"ผู้ที่มีผิวแห้งหยาบและความกระชับลดลง"}}},{"k":"자글한 잔주름이 고민이신 분","t":{"en":{"e":"Are concerned about fine lines"},"jp":{"e":"細かいしわが気になる方"},"cn":{"e":"为细小皱纹烦恼的人"},"tw":{"e":"為細小皺紋困擾的人"},"th":{"e":"ผู้ที่กังวลเรื่องริ้วรอยเล็กๆ"}}},{"k":"보습/탄력/재생 등 복합 시술을 원하시는 분","t":{"en":{"e":"Are looking for a treatment that addresses hydration, firmness, and skin renewal all at once"},"jp":{"e":"保湿・ハリ・再生など総合的なケアを希望する方"},"cn":{"e":"希望进行保湿/紧致/修复等综合护理的人"},"tw":{"e":"希望進行保濕／彈力／修復等複合療程的人"},"th":{"e":"ผู้ที่ต้องการการดูแลแบบครบทั้งความชุ่มชื้น ความกระชับ และการฟื้นฟู"}}},{"k":"*생성형 AI 모델입니다.","t":{"en":{"e":"*THIS IS A GENERATIVE AI MODEL"},"jp":{"e":"*生成型AIモデルです"},"cn":{"e":"*生成式AI模型"},"tw":{"e":"*生成式AI模型"},"th":{"e":"*เป็นภาพโมเดลที่สร้างจาก AI"}}},{"k":"생성형 AI 모델입니다.","t":{"en":{"e":"THIS IS A GENERATIVE AI MODEL"},"jp":{"e":"生成型AIモデルです"},"cn":{"e":"生成式AI模型"},"tw":{"e":"生成式AI模型"},"th":{"e":"เป็นภาพโมเดลที่สร้างจาก AI"}}},{"k":"275,000원","t":{"en":{"e":"275,000 KRW"},"jp":{"e":"275,000 KRW"},"cn":{"e":"27.5万韩元"},"tw":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"27.5만원","t":{"en":{"e":"275,000 KRW"},"jp":{"e":"275,000 KRW"},"cn":{"e":"27.5万韩元"},"tw":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"27.5만","t":{"en":{"e":"275,000 KRW"},"jp":{"e":"275,000 KRW"},"cn":{"e":"27.5万韩元"},"tw":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"*상담 후 결정","t":{"en":{"e":"*Subject to consultation"},"jp":{"e":"※カウンセリング込み"},"cn":{"e":"需先咨询"},"tw":{"e":"需先諮詢"},"th":{"e":"ปรึกษาก่อน"}}},{"k":"※需先咨询","t":{"en":{"e":"*Consultation required first"},"jp":{"e":"※カウンセリング後にご案内"},"cn":{"e":"※需先咨询"},"tw":{"e":"※需先諮詢"},"th":{"e":"※ต้องปรึกษาก่อน"}}},{"k":"※需先諮詢","t":{"en":{"e":"*Consultation required first"},"jp":{"e":"※カウンセリング後にご案内"},"cn":{"e":"※需先咨询"},"tw":{"e":"※需先諮詢"},"th":{"e":"※ต้องปรึกษาก่อน"}}},{"k":"含税","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"含稅","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"รวมภาษีมูลค่าเพิ่มแล้ว","t":{"en":{"e":"VAT INCLUDED"},"jp":{"e":"VAT込み"},"cn":{"e":"含增值税"},"tw":{"e":"含增值稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"무턱필러 + 슈링크 100샷","t":{"en":{"e":"Chin Filler + Shrink 100 Shots"},"tw":{"e":"下巴填充 + Shurink 100發"}},"c":"필러"},{"k":"무턱필러+슈링크 100샷","t":{"en":{"e":"Chin Filler + SHRINK 100 Shots"},"tw":{"e":"下巴填充 + SHRINK 100發"}},"c":"필러"},{"k":"무턱필러","t":{"en":{"e":"Chin Filler"},"tw":{"e":"下巴填充"}},"c":"필러"},{"k":"#인상개선","t":{"en":{"e":"#ImpressionImprovement"},"tw":{"e":"#小臉線條"}}},{"k":"#V라인얼굴","t":{"en":{"e":"#VLineFace"},"tw":{"e":"#V型下巴"}}},{"k":"#얼굴전체탄력","t":{"en":{"e":"#FullFaceElasticity"},"tw":{"e":"#臉部比例平衡"}}},{"k":"무턱","t":{"en":{"e":"recessed chin"},"tw":{"e":"無下巴"}}},{"k":"돌출입","t":{"en":{"e":"protruding mouth"},"tw":{"e":"下巴前凸"}}},{"k":"뭉툭한턱","t":{"en":{"e":"blunt chin"},"tw":{"e":"厚重下巴"}}},{"k":"하관이 고민이라면?","t":{"en":{"e":"Concerned about your lower face?"},"tw":{"e":"如果你在意下半臉？"}}},{"k":"짧고 볼륨이 부족한 턱 끝에 얼굴 전체적인 밸런스가 맞도록 디자인하여 보다 더 갸름하고 세련된 인상으로 개선","t":{"en":{"e":"Designs the short, volume-deficient chin tip to balance the overall face, improving impression to a slimmer and more refined look"},"tw":{"e":"針對下巴較短或下巴末端飽滿度不足的情況 從整體臉部比例出發 進行精細設計 改善為更加立體、精緻的臉部輪廓"}}},{"k":"턱 아래 늘어진 피부를 탄탄하게 잡아주고 흐릿해진 턱선을 위로 끌어올려 윤곽을 또렷하게 정리하며 하관 라인을 매끄럽고 선명하게 완성","t":{"en":{"e":"Firmly tightens sagging skin under the chin and lifts a blurred jawline upward to define the contour, completing a smooth and clear lower face line"},"tw":{"e":"緊實下巴下方 鬆弛的肌膚 將鬆散的輪廓線 向上拉提 讓線條更加俐落 打造俐落分明的下半臉輪廓"}}},{"k":"무턱 때문에 얼굴이 둥글거나 짧아 보이는 경우","t":{"en":{"e":"Cases where the face looks round or short due to a recessed chin"},"tw":{"e":"因下巴問題 臉型看起來較圓或較短的情況"}}},{"k":"턱선이 흐려서 이중턱이 강조되는 경우","t":{"en":{"e":"Cases where a blurred jawline emphasizes a double chin"},"tw":{"e":"下顎線鬆弛 雙下巴明顯的情況"}}},{"k":"필러만 하면 턱이 부자연스러울까 걱정되는 분","t":{"en":{"e":"Those worried that filler alone might look unnatural"},"tw":{"e":"僅進行下巴填充 擔心臉部比例失衡的人"}}},{"k":"수술 없이 하관 라인 정리를 원하시는 분","t":{"en":{"e":"Those who want to refine the lower face line without surgery"},"tw":{"e":"希望不動手術 整理下顎線條的人"}}},{"k":"대용량 갸름주사","t":{"en":{"e":"Large-Capacity Slimming Injection"},"tw":{"e":"大容量 臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"얼굴지방분해주사","t":{"en":{"e":"Face Fat Dissolve Injection"},"tw":{"e":"臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"갸름주사","t":{"en":{"e":"Slimming Injection"},"tw":{"e":"臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"#숨어있던","t":{"en":{"e":"#Hidden"},"tw":{"e":"#隱藏的"}}},{"k":"#V라인찾기","t":{"en":{"e":"#FindVLine"},"tw":{"e":"#尋找V型臉"}}},{"k":"FACE FAT DISSOLVE","t":{"en":{"e":"FACE FAT DISSOLVE"},"tw":{"e":"FACE FAT DISSOLVE"}}},{"k":"단순한 주사 그 이상, 얼굴라인을 되찾다","t":{"en":{"e":"More than just an injection — reclaim your face line"},"tw":{"e":"不只是打一針，找回清晰的臉部線條"}}},{"k":"이중턱 및 볼살 감소","t":{"en":{"e":"Reduces double chin and cheek fat"},"tw":{"e":"減少雙下巴與臉頰脂肪"}}},{"k":"1:1 맞춤 시술","t":{"en":{"e":"1:1 customized treatment"},"tw":{"e":"1:1 客製化療程"}}},{"k":"조화로운 얼굴","t":{"en":{"e":"Harmonious facial proportions"},"tw":{"e":"協調的臉部比例"}}},{"k":"V라인 얼굴형","t":{"en":{"e":"V-line face shape"},"tw":{"e":"V型臉型"}}},{"k":"단순히 빼는 시대는 끝!","t":{"en":{"e":"The era of simply removing fat is over!"},"tw":{"e":"只追求減少的時代已結束！"}}},{"k":"섬세하고 디테일한 몰딩","t":{"en":{"e":"Delicate and detailed molding"},"tw":{"e":"細緻且講究細節的塑形"}}},{"k":"얼굴형을 고려한","t":{"en":{"e":"Tailored to your face shape"},"tw":{"e":"依臉型設計的"}}},{"k":"1:1 맞춤 디자인","t":{"en":{"e":"1:1 customized design"},"tw":{"e":"1:1 客製化設計"}}},{"k":"미미썸 얼굴지방분해주사는 얼굴형과 지방 분포를 고려한 1:1 디자인 시술로 보다 입체적이고 조화로운 작은 얼굴을 연출합니다.","t":{"en":{"e":"MIMISOME face fat dissolve injection considers face shape and fat distribution with 1:1 designed treatment to create a more three-dimensional, harmonious small face."},"tw":{"e":"MIMISOME 臉部溶脂注射考量臉型與脂肪分布 進行 1:1 客製化設計療程 打造協調精緻的小臉。"}}},{"k":"왜 얼굴지방분해주사인가요?","t":{"en":{"e":"Why face fat dissolve injection?"},"tw":{"e":"為什麼選擇臉部溶脂注射？"}}},{"k":"얼굴지방분해주사인가요?","t":{"en":{"e":"Face fat dissolve injection?"},"tw":{"e":"為什麼選擇臉部溶脂注射？"}}},{"k":"1:1 맞춤 디자인 설계","t":{"en":{"e":"1:1 customized design planning"},"tw":{"e":"1:1 客製化設計規劃"}}},{"k":"약 10분 내외 시술","t":{"en":{"e":"Treatment in approx. 10 minutes"},"tw":{"e":"約10分鐘左右療程"}}},{"k":"멍·붓기 적음","t":{"en":{"e":"Minimal bruising and swelling"},"tw":{"e":"將瘀青與腫脹 降至最低"}}},{"k":"지방세포 파괴 후 자연스럽게 체외 배출","t":{"en":{"e":"Fat cells are destroyed and naturally excreted from the body"},"tw":{"e":"脂肪細胞破壞後 自然排出體外"}}},{"k":"#맞춤디자인","t":{"en":{"e":"#CustomDesign"},"tw":{"e":"#客製化設計"}}},{"k":"#섬세한시술","t":{"en":{"e":"#DelicateTreatment"},"tw":{"e":"#細緻療程"}}},{"k":"#슬림한얼굴라인","t":{"en":{"e":"#SlimFaceLine"},"tw":{"e":"#纖細臉部線條"}}},{"k":"작아진 얼굴, 자연스러운 변화로 시작하세요","t":{"en":{"e":"A smaller face — start with natural change"},"tw":{"e":"更小的臉型，從自然的改變開始"}}},{"k":"3D 볼륨필러 5cc","t":{"en":{"e":"3D Volume Filler 5cc"},"tw":{"e":"3D 立體填充 5cc"}},"c":"필러"},{"k":"3D 볼륨필러","t":{"en":{"e":"3D Volume Filler"},"tw":{"e":"3D 立體填充"}},"c":"필러"},{"k":"3D 볼륨 필러","t":{"en":{"e":"3D Volume Filler"},"tw":{"e":"3D 立體填充"}},"c":"필러"},{"k":"3D VOLUME FILLER","t":{"en":{"e":"3D VOLUME FILLER"},"tw":{"e":"3D VOLUME FILLER"}}},{"k":"미미썸 3D 볼륨 필러","t":{"en":{"e":"MIMISOME 3D Volume Filler"},"tw":{"e":"MIMISOME 3D 立體填充"}}},{"k":"미미썸 3D 볼륨필러","t":{"en":{"e":"MIMISOME 3D Volume Filler"},"tw":{"e":"MIMISOME 3D 立體填充"}}},{"k":"빠른 효과로 입체적인 볼륨 완성","t":{"en":{"e":"Fast results for three-dimensional volume"},"tw":{"e":"快速見效 打造立體飽滿輪廓"}}},{"k":"입체적인 윤곽","t":{"en":{"e":"Three-dimensional contour"},"tw":{"e":"立體輪廓"}}},{"k":"자연스러운 볼륨","t":{"en":{"e":"Natural volume"},"tw":{"e":"自然飽滿"}}},{"k":"균형 잡힌 라인","t":{"en":{"e":"Balanced lines"},"tw":{"e":"平衡線條"}}},{"k":"추천 대상","t":{"en":{"e":"Recommended for"},"tw":{"e":"推薦對象"}}},{"k":"꺼진 부위 볼륨을 자연스럽게 채우고 싶은 분","t":{"en":{"e":"Those who want to naturally fill sunken areas"},"tw":{"e":"希望自然填補凹陷部位者"}}},{"k":"얼굴 윤곽과 입체감을 함께 개선하고 싶은 분","t":{"en":{"e":"Those who want to improve both facial contour and dimension"},"tw":{"e":"想同時改善臉部輪廓與比例者"}}},{"k":"자연스럽게 얼굴 균형을 맞추고 싶은 분","t":{"en":{"e":"Those who want to naturally balance their face"},"tw":{"e":"希望自然調整臉部平衡者"}}},{"k":"안전성","t":{"en":{"e":"Safety"},"tw":{"e":"安全性"}}},{"k":"지속력","t":{"en":{"e":"Longevity / Durability"},"tw":{"e":"持久性"}}},{"k":"부위별 구조를 고려한 3D 볼륨 설계로 꺼진 부위는 채우고, 과한 볼륨은 줄여 자연스럽고 입체적인 얼굴라인을 완성합니다.","t":{"en":{"e":"With 3D volume design considering each area's structure, sunken areas are filled and excessive volume is reduced to complete a natural, three-dimensional face line."},"tw":{"e":"依照部位結構設計的3D填充方案 填補凹陷部位，減少過度填充 打造自然立體的臉部線條。"}}},{"k":"시술시간","t":{"en":{"e":"Treatment time"},"tw":{"e":"療程時間"}}},{"k":"일상생활","t":{"en":{"e":"Daily life"},"tw":{"e":"日常生活"}}},{"k":"시술 후 빠르게 가능","t":{"en":{"e":"Resume quickly after treatment"},"tw":{"e":"術後可快速恢復"}}},{"k":"통증정도","t":{"en":{"e":"Pain level"},"tw":{"e":"疼痛程度"}}},{"k":"미미썸의원만의 특별함","t":{"en":{"e":"What makes MIMISOME special"},"tw":{"e":"MIMISOME的獨特優勢"}}},{"k":"사전상담, 섬세한 검진을 통해 개인 상태에 맞는 맞춤형 시술 진행","t":{"en":{"e":"Through pre-consultation and detailed examination, treatment is tailored to your individual condition"},"tw":{"e":"透過術前諮詢與細緻檢查 規劃符合個人狀況的療程"}}},{"k":"풍부한 경험을 바탕으로 편안하고 정교한 시술 제공","t":{"en":{"e":"Comfortable and precise treatment based on rich experience"},"tw":{"e":"憑藉豐富的臨床經驗 提供舒適且精細的療程"}}},{"k":"개인의 신체특성을 고려한 자연스러운 디자인으로 만족도 향상","t":{"en":{"e":"Improved satisfaction with natural design considering individual body characteristics"},"tw":{"e":"考量個人身體特性 打造自然協調的設計"}}},{"k":"108,900원","t":{"en":{"e":"108,900 KRW"},"tw":{"e":"108,900 韓元"}}},{"k":"330,000원","t":{"en":{"e":"330,000 KRW"},"tw":{"e":"330,000 KRW"}}},{"k":"33만","t":{"en":{"e":"330,000 KRW"},"tw":{"e":"33萬韓元"}}},{"k":"*부가세 포함","t":{"en":{"e":"*VAT included"},"tw":{"e":"※含稅"}}},{"k":"부가세 포함","t":{"en":{"e":"VAT included"},"tw":{"e":"含稅"}}}],"glossary":{"하이톡스":{"en":"Hitox","jp":"ハイトックス","cn":"Hitox","tw":"Hitox","th":"Hitox","vn":"Hitox"},"코어톡스":{"en":"Coretox","jp":"コアトックス","cn":"Coretox","tw":"Coretox","th":"Coretox","vn":"Coretox"},"이노톡스":{"en":"Innotox","jp":"イノトックス","cn":"Innotox","tw":"Innotox","th":"Innotox","vn":"Innotox"},"제오민":{"en":"Xeomin","jp":"ゼオミン","cn":"Xeomin","tw":"Xeomin","th":"Xeomin","vn":"Xeomin"},"디스포트":{"en":"Dysport","jp":"ディスポート","cn":"Dysport","tw":"Dysport","th":"Dysport","vn":"Dysport"},"제테마":{"en":"Jetema","jp":"ジェテマ","cn":"Jetema","tw":"Jetema","th":"Jetema","vn":"Jetema"},"보톡스":{"en":"Botox","jp":"ボトックス","cn":"肉毒素","tw":"肉毒素","th":"โบท็อกซ์","vn":"Botox"},"필러":{"en":"Filler","jp":"フィラー","cn":"填充剂","tw":"填充劑","th":"ฟิลเลอร์","vn":"Filler"},"리프팅":{"en":"Lifting","jp":"リフトアップ","cn":"提升","tw":"提升","th":"ยกกระชับ","vn":"Nâng cơ"},"스킨부스터":{"en":"Skin Booster","jp":"スキンブースター","cn":"皮肤增强剂","tw":"皮膚增強劑","th":"สกินบูสเตอร์","vn":"Skin Booster"},"실리프팅":{"en":"Thread Lifting","jp":"糸リフト","cn":"线雕","tw":"線雕","th":"ร้อยไหม","vn":"Căng chỉ"},"윤곽주사":{"en":"Face Contouring Injection","jp":"輪郭注射","cn":"轮廓针","tw":"輪廓針","th":"ฉีดสลายไขมัน","vn":"Tiêm thon gọn"},"스킨보톡스":{"en":"Skin Botox","jp":"スキンボトックス","cn":"皮肤肉毒素","tw":"皮膚肉毒素","th":"สกินโบท็อกซ์","vn":"Skin Botox"},"사각턱":{"en":"Jawline / Masseter","jp":"エラ","cn":"方下巴","tw":"方下巴","th":"กราม","vn":"Hàm vuông"},"쥬베룩":{"en":"Juvelook","jp":"ジュベルック","cn":"Juvelook","tw":"Juvelook","th":"Juvelook","vn":"Juvelook"},"리쥬란":{"en":"Rejuran","jp":"リジュラン","cn":"Rejuran","tw":"Rejuran","th":"Rejuran","vn":"Rejuran"},"스킨부스터 리투오":{"en":"Rituo Skin Booster","jp":"リトゥオスキンブースター","cn":"Rituo 皮肤增强剂","tw":"Rituo 皮膚增強劑","th":"Rituo Skin Booster","vn":"Rituo Skin Booster"},"아쿠아필":{"en":"Aqua Peel","jp":"アクアピール","cn":"水光焕肤","tw":"水光煥膚","th":"Aqua Peel","vn":"Aqua Peel"},"아포지":{"en":"Apogee","jp":"アポジー","cn":"Apogee","tw":"Apogee","th":"Apogee","vn":"Apogee"},"피코토닝":{"en":"Pico Toning","jp":"ピコトーニング","cn":"皮秒调理","tw":"皮秒調理","th":"Pico Toning","vn":"Pico Toning"},"이마":{"en":"Forehead","jp":"おでこ","cn":"额头","tw":"額頭","th":"หน้าผาก","vn":"trán"},"미간":{"en":"Frown Lines","jp":"眉間","cn":"眉间","tw":"眉間","th":"ระหว่างคิ้ว","vn":"giữa chân mày"},"눈가":{"en":"Crow's Feet","jp":"目元","cn":"眼周","tw":"眼周","th":"หางตา","vn":"đuôi mắt"},"콧대":{"en":"Nasal Bridge","jp":"鼻柱","cn":"鼻梁","tw":"鼻梁","th":"สันจมูก","vn":"sống mũi"},"콧등":{"en":"Bunny Lines","jp":"鼻背","cn":"鼻背","tw":"鼻背","th":"ดั้งจมูก","vn":"lưng mũi"},"자갈턱":{"en":"Chin","jp":"顎の梅干しじわ","cn":"酒窝下巴","tw":"酒窩下巴","th":"คาง","vn":"cằm đá cuội"},"국산":{"en":"Korean","jp":"韓国製","cn":"韩国产","tw":"韓國產","th":"เกาหลี","vn":"nội địa"},"수입":{"en":"Imported","jp":"海外製","cn":"进口","tw":"進口","th":"นำเข้า","vn":"nhập khẩu"},"1부위":{"en":"1 Area","jp":"1部位","cn":"1部位","tw":"1部位","th":"1 บริเวณ","vn":"1 vùng"},"풀페이스":{"en":"Full Face","jp":"フルフェイス","cn":"全脸","tw":"全臉","th":"ทั่วหน้า","vn":"Toàn mặt"},"1회":{"en":"1 Session","jp":"1回","cn":"1次","tw":"1次","th":"1 ครั้ง","vn":"1 lần"},"체험가":{"en":"Trial Price","jp":"体験価格","cn":"体验价","tw":"體驗價","th":"ราคาทดลอง","vn":"Giá trải nghiệm"},"첫방문":{"en":"First Visit","jp":"初回来院","cn":"首次到访","tw":"首次到訪","th":"เยี่ยมชมครั้งแรก","vn":"Lần đầu đến"},"남자":{"en":"Male","jp":"男性","cn":"男士","tw":"男士","th":"ผู้ชาย","vn":"Nam"},"여자":{"en":"Female","jp":"女性","cn":"女士","tw":"女士","th":"ผู้หญิง","vn":"Nữ"},"택1":{"en":"Choose 1","jp":"1選択","cn":"选1","tw":"選1","th":"เลือก 1","vn":"Chọn 1"},"미미썸":{"en":"Mimisome","jp":"ミミソム","cn":"美美썸","tw":"MIMISOME","th":"Mimisome","vn":"Mimisome"},"미미썸 한정":{"en":"Mimisome Limited","jp":"ミミソム限定","cn":"MIMISOME 限定","tw":"MIMISOME 限定","th":"Mimisome Limited","vn":"Ưu đãi Mimisome"},"한정":{"en":"Limited","jp":"限定","cn":"限定","tw":"限定","th":"Limited","vn":"Giới hạn"},"VAT포함":{"en":"VAT included","jp":"VAT込み","cn":"含增值税","tw":"含增值稅","th":"รวม VAT","vn":"Đã bao gồm VAT"},"메조주사":{"en":"Mesotherapy Injection","jp":"メソセラピー注射","cn":"中胚层疗法注射","tw":"中胚層療法注射","th":"Mesotherapy","vn":"Tiêm mesotherapy"},"메조테라피":{"en":"Mesotherapy","jp":"メソセラピー","cn":"中胚层疗法","tw":"中胚層療法","th":"Mesotherapy","vn":"Mesotherapy"},"두피":{"en":"scalp","jp":"頭皮","cn":"头皮","tw":"頭皮","th":"หนังศีรษะ","vn":"da đầu"},"모근":{"en":"hair root","jp":"毛根","cn":"毛根","tw":"毛根","th":"รากผม","vn":"chân tóc"},"모낭":{"en":"hair follicle","jp":"毛包","cn":"毛囊","tw":"毛囊","th":"รูขุมขน","vn":"nang tóc"},"탈모":{"en":"hair loss","jp":"脱毛/発毛","cn":"脱发","tw":"落髮","th":"ผมร่วง","vn":"rụng tóc"},"탈모치료":{"en":"Hair Loss Treatment","jp":"発毛治療","cn":"脱发治疗","tw":"落髮治療","th":"รักษาผมร่วง","vn":"Điều trị rụng tóc"},"약물요법":{"en":"medication therapy","jp":"薬物療法","cn":"药物疗法","tw":"藥物療法","th":"การรักษาด้วยยา","vn":"liệu pháp dùng thuốc"},"약물처방":{"en":"Medication Prescription","jp":"薬処方","cn":"药物处方","tw":"藥物處方","th":"ยาที่สั่งโดยแพทย์","vn":"kê đơn thuốc"},"광에너지":{"en":"light energy","jp":"光エネルギー","cn":"光能","tw":"光能","th":"พลังงานแสง","vn":"năng lượng ánh sáng"},"테라피":{"en":"therapy","jp":"療法","cn":"疗法","tw":"療法","th":"บำบัด","vn":"liệu pháp"},"시술":{"en":"procedure","jp":"施術","cn":"治疗","tw":"療程","th":"การรักษา","vn":"thủ thuật"},"마취":{"en":"anesthesia","jp":"麻酔","cn":"麻醉","tw":"麻醉","th":"ยาชา","vn":"gây tê"},"연고 마취":{"en":"topical anesthesia","jp":"麻酔クリーム","cn":"表面麻醉","tw":"表面麻醉","th":"ยาชาทา","vn":"gây tê bề mặt"},"일상생활":{"en":"daily life","jp":"日常生活","cn":"日常生活","tw":"日常生活","th":"ชีวิตประจำวัน","vn":"sinh hoạt hàng ngày"},"복귀":{"en":"return","jp":"復帰","cn":"恢复","tw":"恢復","th":"กลับสู่","vn":"trở lại"},"추천 주기":{"en":"recommended cycle","jp":"推奨周期","cn":"推荐周期","tw":"建議週期","th":"ความถี่แนะนำ","vn":"tần suất khuyến nghị"},"맞춤 상담":{"en":"customized consultation","jp":"個別カウンセリング","cn":"定制咨询","tw":"客製化諮詢","th":"ปรึกษาส่วนบุคคล","vn":"tư vấn cá nhân hoá"},"섬유아세포":{"en":"fibroblast","jp":"線維芽細胞","cn":"纤维母细胞","tw":"纖維母細胞","th":"ไฟโบรบลาสต์","vn":"tế bào sợi"},"재생":{"en":"regeneration","jp":"再生","cn":"再生","tw":"再生","th":"ฟื้นฟู","vn":"tái tạo"},"산후":{"en":"postpartum","jp":"産後","cn":"产后","tw":"產後","th":"หลังคลอด","vn":"sau sinh"},"은밀 부위":{"en":"intimate areas","jp":"デリケートゾーン","cn":"私密部位","tw":"私密部位","th":"จุดซ่อนเร้น","vn":"vùng kín"},"여성 은밀 부위":{"en":"women's intimate areas","jp":"女性デリケートゾーン","cn":"女性私密部位","tw":"女性私密部位","th":"จุดซ่อนเร้นสำหรับผู้หญิง","vn":"vùng kín của phụ nữ"},"질":{"en":"vaginal region","jp":"膣","cn":"阴道","tw":"陰道","th":"บริเวณช่องคลอด","vn":"âm đạo"},"탄력":{"en":"elasticity","jp":"弾力","cn":"弹性","tw":"彈性","th":"ความยืดหยุ่น","vn":"độ đàn hồi"},"볼륨":{"en":"volume","jp":"ボリューム","cn":"丰盈","tw":"豐盈","th":"วอลลุ่ม","vn":"thể tích"},"볼륨감":{"en":"volume","jp":"ボリューム感","cn":"饱满感","tw":"豐盈感","th":"ความฟูดูเป็นธรรมชาติ","vn":"cảm giác đầy đặn"},"자신감":{"en":"confidence","jp":"自信","cn":"自信","tw":"自信","th":"ความมั่นใจ","vn":"sự tự tin"},"출산":{"en":"childbirth","jp":"出産","cn":"分娩","tw":"生產","th":"การคลอด","vn":"sinh nở"},"노화":{"en":"aging","jp":"加齢","cn":"老化","tw":"老化","th":"อายุที่เพิ่มขึ้น","vn":"lão hóa"},"임상 경험":{"en":"clinical experience","jp":"臨床経験","cn":"临床经验","tw":"臨床經驗","th":"ประสบการณ์ทางคลินิก","vn":"kinh nghiệm lâm sàng"},"통증":{"en":"pain","jp":"痛み","cn":"疼痛","tw":"疼痛","th":"ความเจ็บ","vn":"đau"},"회복":{"en":"recovery","jp":"回復","cn":"恢复","tw":"恢復","th":"การฟื้นตัว","vn":"phục hồi"},"1:1 맞춤":{"en":"1:1 personalized","jp":"1:1 オーダーメイド","cn":"1:1 定制","tw":"1:1 客製化","th":"ตัวต่อตัว","vn":"1:1 cá nhân hóa"},"지속성":{"en":"longevity","jp":"持続性","cn":"持久性","tw":"持久性","th":"ความคงทน","vn":"tính bền lâu"},"안전성":{"en":"safety","jp":"安全性","cn":"安全性","tw":"安全性","th":"ความปลอดภัย","vn":"tính an toàn"},"자연스러운":{"en":"natural","jp":"自然な","cn":"自然","tw":"自然","th":"เป็นธรรมชาติ","vn":"tự nhiên"},"필러 성분":{"en":"filler ingredients","jp":"フィラー成分","cn":"填充成分","tw":"填充成分","th":"ส่วนผสมฟิลเลอร์","vn":"thành phần filler"},"히알루론산":{"en":"hyaluronic acid","jp":"ヒアルロン酸","cn":"玻尿酸","tw":"玻尿酸","th":"กรดไฮยาลูโรนิก","vn":"axit hyaluronic"},"인티메이트":{"en":"intimate","jp":"インティメート","cn":"私密","tw":"親密","th":"อินทิเมท","vn":"vùng kín"},"여성 인티메이트":{"en":"women's intimate area","jp":"女性インティメート","cn":"女性私密","tw":"女性親密部位","th":"อินทิเมทเฉพาะผู้หญิง","vn":"vùng kín nữ giới"},"성감":{"en":"sensitivity","jp":"感度","cn":"敏感度","tw":"敏感度","th":"ความรู้สึก","vn":"độ nhạy cảm"},"지속력":{"en":"longevity","jp":"持続力","cn":"持久力","tw":"持久力","th":"ความคงทน","vn":"tính bền lâu"},"맞춤상담":{"en":"personalized consultation","jp":"オーダーメイドカウンセリング","cn":"定制咨询","tw":"客製化諮詢","th":"ปรึกษาเฉพาะบุคคล","vn":"tư vấn cá nhân hóa"},"심미성":{"en":"aesthetics","jp":"審美性","cn":"审美","tw":"美感","th":"ความสวยงาม","vn":"thẩm mỹ"},"기능성":{"en":"functionality","jp":"機能性","cn":"功能性","tw":"功能性","th":"การใช้งาน","vn":"tính năng"},"의료진":{"en":"medical staff","jp":"医療スタッフ","cn":"医疗团队","tw":"醫療團隊","th":"ทีมแพทย์","vn":"đội ngũ y tế"},"노하우":{"en":"expertise","jp":"ノウハウ","cn":"经验技术","tw":"經驗技術","th":"ความเชี่ยวชาญ","vn":"bí quyết"},"미미썸의원":{"en":"MIMISOME Clinic","jp":"ミミソム医院","cn":"MIMISOME诊所","tw":"MIMISOME診所","th":"คลินิก MIMISOME","vn":"Phòng khám MIMISOME"},"엑소좀":{"en":"exosome","jp":"エクソソーム","cn":"外泌体","tw":"外泌體","th":"เอ็กโซโซม","vn":"exosome"},"ASCE+":{"en":"ASCE+","jp":"ASCE+","cn":"ASCE+","tw":"ASCE+","th":"ASCE+","vn":"ASCE+"},"콜라겐":{"en":"collagen","jp":"コラーゲン","cn":"胶原蛋白","tw":"膠原蛋白","th":"คอลลาเจน","vn":"collagen"},"엘라스틴":{"en":"elastin","jp":"エラスチン","cn":"弹性蛋白","tw":"彈力蛋白","th":"อีลาสติน","vn":"elastin"},"모공":{"en":"pore","jp":"毛穴","cn":"毛孔","tw":"毛孔","th":"รูขุมขน","vn":"lỗ chân lông"},"모공 축소":{"en":"pore care","jp":"毛穴ケア","cn":"毛孔缩小","tw":"毛孔護理","th":"ดูแลรูขุมขน","vn":"thu nhỏ lỗ chân lông"},"모공지우개":{"en":"Pore Care","jp":"毛穴ケア","cn":"毛孔护理","tw":"毛孔護理","th":"ลบเลือนรูขุมขน","vn":"xóa lỗ chân lông"},"피부결":{"en":"skin texture","jp":"肌キメ","cn":"肤质","tw":"膚質","th":"ผิวสัมผัส","vn":"kết cấu da"},"유수분":{"en":"moisture and oil balance","jp":"水分油分","cn":"水油平衡","tw":"水油平衡","th":"ความชุ่มชื้นและน้ำมัน","vn":"cân bằng dầu và nước"},"유수분 밸런스":{"en":"Hydration Balance","jp":"水分バランス","cn":"水油平衡","tw":"水分平衡","th":"สมดุลความชุ่มชื้น","vn":"cân bằng độ ẩm"},"트러블":{"en":"skin problems / acne","jp":"肌トラブル","cn":"肌肤问题","tw":"肌膚問題","th":"ปัญหาผิว","vn":"vấn đề da"},"침습적":{"en":"invasive","jp":"侵襲的","cn":"侵入式","tw":"侵入式","th":"แบบรุกราน","vn":"xâm lấn"},"비침습":{"en":"non-invasive","jp":"非侵襲","cn":"非侵入式","tw":"非侵入式","th":"ไม่รุกราน","vn":"không xâm lấn"},"근본":{"en":"fundamental / root cause","jp":"根本","cn":"根本","tw":"根本","th":"พื้นฐาน","vn":"căn nguyên"},"광택":{"en":"glow","jp":"ツヤ","cn":"光泽","tw":"光澤","th":"ความเงา","vn":"độ bóng"},"팅커벨":{"en":"Tinker Bell","jp":"ティンカーベル","cn":"小精灵","tw":"精靈","th":"ทิงเกอร์เบลล์","vn":"Tinker Bell"},"귀":{"en":"ear","jp":"耳","cn":"耳朵","tw":"耳朵","th":"หู","vn":"tai"},"귓볼":{"en":"earlobe","jp":"耳たぶ","cn":"耳垂","tw":"耳垂","th":"ติ่งหู","vn":"dái tai"},"귀필러":{"en":"ear filler","jp":"耳フィラー","cn":"耳部填充","tw":"精靈耳玻尿酸","th":"ฟิลเลอร์หู","vn":"filler tai"},"소얼굴":{"en":"small face / face slimming","jp":"小顔","cn":"小脸","tw":"小顏","th":"หน้าเล็ก","vn":"mặt nhỏ"},"얼굴 비율":{"en":"facial proportions","jp":"顔のバランス","cn":"脸部比例","tw":"臉部比例","th":"สัดส่วนใบหน้า","vn":"tỉ lệ khuôn mặt"},"이목구비":{"en":"facial features","jp":"顔のパーツ","cn":"五官","tw":"五官","th":"ใบหน้า","vn":"ngũ quan"},"비대칭":{"en":"asymmetrical","jp":"非対称","cn":"不对称","tw":"不對稱","th":"ไม่สมมาตร","vn":"không đối xứng"},"이상적인":{"en":"ideal","jp":"理想的な","cn":"理想的","tw":"理想的","th":"ที่เหมาะสมที่สุด","vn":"lý tưởng"},"비율":{"en":"proportions","jp":"比率","cn":"比例","tw":"比例","th":"สัดส่วน","vn":"tỉ lệ"},"균형":{"en":"balance","jp":"バランス","cn":"平衡","tw":"平衡","th":"สมดุล","vn":"cân bằng"},"조화로운":{"en":"harmonious / balanced","jp":"調和のとれた","cn":"协调","tw":"協調","th":"กลมกลืน","vn":"hài hòa"},"비결":{"en":"secret / tip","jp":"秘訣","cn":"秘诀","tw":"秘訣","th":"เคล็ดลับ","vn":"bí quyết"},"얇은":{"en":"thin","jp":"薄い","cn":"薄","tw":"薄","th":"บาง","vn":"mỏng"},"정면":{"en":"front view","jp":"正面","cn":"正面","tw":"正面","th":"ด้านหน้า","vn":"mặt trước"},"써마지":{"en":"Thermage","jp":"サーマクール","cn":"Thermage","tw":"鳳凰電波","th":"เทอร์มาจ","vn":"Thermage"},"써마지FLX":{"en":"Thermage FLX","jp":"サーマクールFLX","cn":"Thermage FLX","tw":"第四代鳳凰電波","th":"เทอร์มาจ FLX","vn":"Thermage FLX"},"써마지 FLX":{"en":"Thermage FLX","jp":"サーマクールFLX","cn":"Thermage FLX","tw":"第四代鳳凰電波","th":"เทอร์มาจ FLX","vn":"Thermage FLX"},"FLX":{"en":"FLX","jp":"FLX","cn":"FLX","tw":"FLX","th":"FLX","vn":"FLX"},"고주파":{"en":"radiofrequency / RF","jp":"高周波","cn":"高频","tw":"電波","th":"คลื่นวิทยุ","vn":"sóng RF"},"고주파 에너지":{"en":"radiofrequency energy","jp":"高周波エネルギー","cn":"射频能量","tw":"電波能量","th":"พลังงานคลื่นวิทยุ","vn":"năng lượng sóng RF"},"잔주름":{"en":"fine lines","jp":"小ジワ","cn":"细纹","tw":"細紋","th":"ริ้วรอยเล็กๆ","vn":"nếp nhăn nhỏ"},"윤곽":{"en":"contour","jp":"輪郭","cn":"轮廓","tw":"輪廓","th":"โครงหน้า","vn":"đường viền khuôn mặt"},"타이트닝":{"en":"tightening","jp":"タイトニング","cn":"紧致","tw":"緊實","th":"กระชับ","vn":"săn chắc"},"스킨 타이트닝":{"en":"skin tightening","jp":"スキンタイトニング","cn":"皮肤紧致","tw":"肌膚緊實","th":"กระชับผิว","vn":"săn chắc da"},"300샷":{"en":"300 Shots","jp":"300ショット","cn":"300发","tw":"300發","th":"300 ช็อต","vn":"300 shot"},"샷":{"en":"shot","jp":"ショット","cn":"发","tw":"發","th":"ช็อต","vn":"shot"},"처짐":{"en":"sagging","jp":"たるみ","cn":"松弛","tw":"鬆弛","th":"หย่อนคล้อย","vn":"chảy xệ"},"처진":{"en":"sagging","jp":"たるんだ","cn":"松弛的","tw":"鬆弛的","th":"หย่อนคล้อย","vn":"chảy xệ"},"턱선":{"en":"jawline","jp":"フェイスライン","cn":"下颌线","tw":"下顎線","th":"กรอบหน้า","vn":"đường viền hàm"},"얼굴라인":{"en":"facial contours","jp":"フェイスライン","cn":"脸部线条","tw":"臉部輪廓","th":"โครงหน้า","vn":"đường nét khuôn mặt"},"입가":{"en":"around the mouth","jp":"口元","cn":"嘴角","tw":"嘴角","th":"รอบปาก","vn":"vùng quanh miệng"},"주름":{"en":"wrinkle","jp":"シワ","cn":"皱纹","tw":"皺紋","th":"ริ้วรอย","vn":"nếp nhăn"},"진동":{"en":"vibration","jp":"振動","cn":"震动","tw":"震動","th":"การสั่นสะเทือน","vn":"rung"},"전방향":{"en":"multidirectional","jp":"全方向","cn":"全方位","tw":"全方位","th":"รอบทิศทาง","vn":"đa hướng"},"맞춤":{"en":"customized","jp":"オーダーメイド","cn":"定制","tw":"客製化","th":"ปรับเฉพาะบุคคล","vn":"cá nhân hoá"},"광고용 이미지":{"en":"advertising image","jp":"広告用イメージ","cn":"广告形象","tw":"廣告形象","th":"ภาพโฆษณา","vn":"hình ảnh quảng cáo"},"포텐자":{"en":"Potenza","jp":"ポテンツァ","cn":"Potenza","tw":"Potenza","th":"โพเทนซ่า","vn":"Potenza"},"포텐자 레이저":{"en":"Potenza Laser","jp":"ポテンツァレーザー","cn":"Potenza激光","tw":"Potenza雷射","th":"โพเทนซ่าเลเซอร์","vn":"Potenza Laser"},"레이저":{"en":"laser","jp":"レーザー","cn":"激光","tw":"雷射","th":"เลเซอร์","vn":"laser"},"모공축소":{"en":"pore reduction","jp":"毛穴縮小","cn":"缩小毛孔","tw":"縮小毛孔","th":"ลดรูขุมขน","vn":"thu nhỏ lỗ chân lông"},"흉터":{"en":"scar","jp":"傷跡","cn":"疤痕","tw":"疤痕","th":"แผลเป็น","vn":"sẹo"},"흉터치료":{"en":"scar treatment","jp":"傷跡治療","cn":"疤痕治疗","tw":"疤痕治療","th":"รักษาแผลเป็น","vn":"điều trị sẹo"},"여드름":{"en":"acne","jp":"ニキビ","cn":"痤疮","tw":"痘痘","th":"สิว","vn":"mụn trứng cá"},"여드름 흉터":{"en":"acne scar","jp":"ニキビ跡","cn":"痘印 / 痘疤","tw":"痘疤","th":"แผลสิว","vn":"sẹo mụn"},"탄력개선":{"en":"elasticity improvement","jp":"ハリ改善","cn":"弹性改善","tw":"彈力改善","th":"เพิ่มความกระชับ","vn":"cải thiện độ đàn hồi"},"깐달걀피부":{"en":"glass skin","jp":"ゆで卵肌","cn":"剥壳鸡蛋肌","tw":"剝殼雞蛋肌","th":"ผิวไข่ต้ม","vn":"da trứng gà bóc vỏ"},"자신감UP":{"en":"confidence UP","jp":"自信アップ","cn":"自信UP","tw":"自信UP","th":"เพิ่มความมั่นใจ","vn":"tự tin UP"},"피부층":{"en":"skin layer","jp":"皮膚層","cn":"皮肤层","tw":"皮膚層","th":"ชั้นผิว","vn":"lớp da"},"깊이":{"en":"depth","jp":"深さ","cn":"深度","tw":"深度","th":"ความลึก","vn":"độ sâu"},"종류":{"en":"type","jp":"種類","cn":"种类","tw":"種類","th":"ชนิด","vn":"loại"},"에너지":{"en":"energy","jp":"エネルギー","cn":"能量","tw":"能量","th":"พลังงาน","vn":"năng lượng"},"시술횟수":{"en":"number of sessions","jp":"施術回数","cn":"治疗次数","tw":"療程次數","th":"จำนวนครั้งที่ทำ","vn":"số lần điều trị"},"키메 / 피부결":{"en":"skin texture","jp":"キメ","cn":"肤质","tw":"膚質","th":"ผิวสัมผัส","vn":"kết cấu da"},"RF":{"en":"RF","jp":"RF / 高周波","cn":"射频","tw":"電波","th":"RF","vn":"sóng RF"},"카운셀링":{"en":"consultation","jp":"カウンセリング","cn":"咨询","tw":"諮詢","th":"คำปรึกษา","vn":"tư vấn"},"카운셀링 포함":{"en":"consultation included","jp":"カウンセリング込み","cn":"含咨询","tw":"含諮詢","th":"รวมค่าปรึกษา","vn":"đã bao gồm tư vấn"},"리프트업":{"en":"lift-up","jp":"リフトアップ","cn":"提升","tw":"提拉","th":"ยกกระชับ","vn":"nâng cơ"},"예민":{"en":"sensitive","jp":"敏感","cn":"敏感","tw":"敏感","th":"แพ้ง่าย","vn":"nhạy cảm"},"리쥬란HB":{"en":"REJURAN HB","jp":"リジュランHB","cn":"丽珠兰 HB","tw":"麗珠蘭 HB","th":"REJURAN HB","vn":"Rejuran HB"},"리쥬란HB PLUS":{"en":"REJURAN HB PLUS","jp":"リジュラン HB PLUS","cn":"丽珠兰 HB PLUS","tw":"麗珠蘭 HB PLUS","th":"REJURAN HB PLUS","vn":"Rejuran HB PLUS"},"리쥬란 힐러":{"en":"Rejuran Healer","jp":"リジュラン ヒーラー / リジュラン成分","cn":"丽珠兰成分","tw":"麗珠蘭成分","th":"สารรีจูรัน","vn":"Rejuran Healer"},"HB PLUS":{"en":"HB PLUS","jp":"HB PLUS","cn":"HB PLUS","tw":"HB PLUS","th":"HB PLUS","vn":"HB PLUS"},"안티에이징":{"en":"anti-aging","jp":"アンチエイジング","cn":"抗衰老","tw":"抗老化","th":"การชะลอวัย","vn":"chống lão hóa"},"연어주사":{"en":"salmon injection","jp":"サーモン注射","cn":"鲑鱼针","tw":"鮭魚針","th":"เข็มแซลมอน","vn":"tiêm cá hồi"},"연어":{"en":"salmon","jp":"サーモン","cn":"鲑鱼","tw":"鮭魚","th":"แซลมอน","vn":"cá hồi"},"폴리뉴클레오타이드":{"en":"polynucleotide / PN","jp":"ポリヌクレオチド","cn":"多核苷酸","tw":"多核苷酸","th":"โพลีนิวคลีโอไทด์","vn":"polynucleotide"},"PN성분":{"en":"PN ingredient","jp":"PN成分","cn":"PN成分","tw":"PN成分","th":"สาร PN","vn":"thành phần PN"},"리도카인":{"en":"lidocaine","jp":"リドカイン","cn":"利多卡因","tw":"利多卡因","th":"ลิโดเคน","vn":"lidocaine"},"국소마취":{"en":"local anesthesia","jp":"局所麻酔","cn":"局部麻醉","tw":"局部麻醉","th":"ยาชาเฉพาะที่","vn":"gây tê tại chỗ"},"마취성분":{"en":"anesthetic ingredient","jp":"麻酔成分","cn":"麻醉成分","tw":"麻醉成分","th":"ส่วนผสมยาชา","vn":"thành phần gây tê"},"수분감":{"en":"hydration","jp":"保湿感","cn":"水润感","tw":"保濕感","th":"ความชุ่มชื้น","vn":"độ ẩm"},"보습":{"en":"moisturizing / hydration","jp":"保湿","cn":"保湿","tw":"保濕","th":"ให้ความชุ่มชื้น","vn":"dưỡng ẩm"},"피부재생":{"en":"skin regeneration","jp":"皮膚再生","cn":"皮肤再生","tw":"皮膚再生","th":"การฟื้นฟูผิว","vn":"tái tạo da"},"탄력상승":{"en":"improved elasticity","jp":"弾力アップ","cn":"弹力提升","tw":"彈力提升","th":"เพิ่มความกระชับ","vn":"tăng độ đàn hồi"},"피부결 개선":{"en":"improved skin texture","jp":"肌キメ改善","cn":"肤质改善","tw":"膚質改善","th":"ผิวเรียบเนียนขึ้น","vn":"cải thiện kết cấu da"},"윤기":{"en":"radiance / glow","jp":"ツヤ","cn":"光泽","tw":"光澤","th":"ความเปล่งปลั่ง","vn":"độ bóng / sáng"},"생기":{"en":"vitality","jp":"ハリ / 生気","cn":"活力","tw":"活力","th":"ความสดใส","vn":"sức sống"},"푸석함":{"en":"dryness / rough","jp":"カサつき","cn":"粗糙","tw":"粗糙","th":"หยาบกร้าน","vn":"thô ráp"},"건조함":{"en":"dryness","jp":"乾燥","cn":"干燥","tw":"乾燥","th":"แห้ง","vn":"khô"},"분해":{"en":"break down / decompose","jp":"分解","cn":"分解","tw":"分解","th":"สลายตัว","vn":"phân hủy"},"형태 유지":{"en":"maintains structure","jp":"形状を保つ","cn":"保持形态","tw":"保持形態","th":"คงรูปทรง","vn":"giữ hình dạng"},"생체 적합":{"en":"biocompatible","jp":"生体適合","cn":"生物相容","tw":"生物相容","th":"เข้ากันได้กับร่างกาย","vn":"tương thích sinh học"},"쥬베룩 스킨":{"en":"Juvelook Skin","jp":"ジュベルックスキン","cn":"乔雅露","tw":"喬雅露","th":"Juvelook","vn":"Juvelook Skin"},"JUVELOOK":{"en":"JUVELOOK","jp":"ジュベルック","cn":"乔雅露","tw":"喬雅露","th":"JUVELOOK","vn":"Juvelook"},"콜라겐 생성":{"en":"collagen production","jp":"コラーゲン生成","cn":"胶原蛋白生成","tw":"膠原蛋白生成","th":"การสร้างคอลลาเจน","vn":"sản sinh collagen"},"자가콜라겐":{"en":"natural / self collagen","jp":"自己コラーゲン","cn":"自身胶原","tw":"自身膠原","th":"คอลลาเจนตามธรรมชาติ","vn":"collagen tự thân"},"물광":{"en":"radiant glow / dewy","jp":"ツヤ / 水光","cn":"水光","tw":"水光","th":"ผิวฉ่ำโกลว์","vn":"da căng bóng"},"물광 피부":{"en":"glass-like dewy skin","jp":"水光肌","cn":"水光肌","tw":"水光肌","th":"ผิวฉ่ำวาว","vn":"da căng bóng"},"긴장":{"en":"tightening","jp":"引き締め","cn":"紧致","tw":"緊緻","th":"กระชับ","vn":"săn chắc"},"PLA":{"en":"PLA (polylactic acid)","jp":"PLA","cn":"PLA","tw":"PLA","th":"PLA","vn":"PLA"},"HA":{"en":"HA (hyaluronic acid)","jp":"HA","cn":"HA","tw":"HA","th":"HA","vn":"HA"},"Hybrid filler":{"en":"Hybrid filler","jp":"ハイブリッドフィラー","cn":"混合填充剂","tw":"混合填充劑","th":"ฟิลเลอร์ไฮบริด","vn":"filler hybrid"},"거칠다":{"en":"rough","jp":"ごわつく","cn":"粗糙","tw":"粗糙","th":"หยาบ","vn":"thô ráp"},"자글한":{"en":"fine / wrinkly","jp":"細かい","cn":"细小","tw":"細小","th":"เล็กๆ","vn":"nhỏ li ti"},"복합 시술":{"en":"combination treatment","jp":"総合的なケア","cn":"综合护理","tw":"複合療程","th":"การดูแลแบบครบ","vn":"liệu trình kết hợp"},"AI 모델":{"en":"AI model","jp":"AIモデル","cn":"AI模型","tw":"AI模型","th":"โมเดล AI","vn":"mô hình AI"},"생성형 AI":{"en":"generative AI","jp":"生成型AI","cn":"生成式AI","tw":"生成式AI","th":"AI สร้างภาพ","vn":"AI tạo sinh"},"무턱":{"en":"recessed chin","jp":"後退顎","cn":"无下巴","tw":"無下巴","th":"คางสั้น","vn":"cằm lẹm"},"무턱필러":{"en":"Chin Filler","jp":"無顎フィラー","cn":"无下巴填充","tw":"下巴填充","th":"ฟิลเลอร์คาง","vn":"filler cằm lẹm"},"돌출입":{"en":"protruding mouth","jp":"突き出た口","cn":"突嘴","tw":"下巴前凸","th":"ปากยื่น","vn":"miệng vẩu"},"이중턱":{"en":"double chin","jp":"二重あご","cn":"双下巴","tw":"雙下巴","th":"เหนียงคู่","vn":"cằm đôi"},"하관":{"en":"lower face","jp":"下顔","cn":"下半脸","tw":"下半臉","th":"ใบหน้าส่วนล่าง","vn":"phần dưới khuôn mặt"},"갸름주사":{"en":"slimming injection","jp":"スリミング注射","cn":"瘦脸针","tw":"溶脂注射","th":"ฉีดเรียวหน้า","vn":"tiêm thon gọn"},"얼굴지방분해주사":{"en":"Face Fat Dissolve Injection","jp":"顔脂肪溶解注射","cn":"面部溶脂针","tw":"臉部溶脂注射","th":"ฉีดสลายไขมันใบหน้า","vn":"tiêm tan mỡ mặt"},"지방분해주사":{"en":"fat dissolve injection","jp":"脂肪溶解注射","cn":"溶脂针","tw":"溶脂注射","th":"ฉีดสลายไขมัน","vn":"tiêm tan mỡ"},"볼살":{"en":"cheek fat","jp":"頬の肉","cn":"颊脂肪","tw":"臉頰脂肪","th":"ไขมันแก้ม","vn":"mỡ má"},"V라인":{"en":"V-line","jp":"Vライン","cn":"V线","tw":"V型","th":"V-line","vn":"V-line"},"지방세포":{"en":"fat cell","jp":"脂肪細胞","cn":"脂肪细胞","tw":"脂肪細胞","th":"เซลล์ไขมัน","vn":"tế bào mỡ"},"붓기":{"en":"swelling","jp":"腫れ","cn":"肿胀","tw":"腫脹","th":"อาการบวม","vn":"sưng"},"멍":{"en":"bruising","jp":"あざ","cn":"瘀青","tw":"瘀青","th":"รอยช้ำ","vn":"vết bầm"},"3D 볼륨필러":{"en":"3D Volume Filler","jp":"3Dボリュームフィラー","cn":"3D立体填充","tw":"3D 立體填充","th":"3D Volume Filler","vn":"Filler 3D"},"볼륨필러":{"en":"Volume Filler","jp":"ボリュームフィラー","cn":"立体填充","tw":"立體填充","th":"Volume Filler","vn":"Volume Filler"},"꺼진 부위":{"en":"sunken area","jp":"くぼんだ部分","cn":"凹陷部位","tw":"凹陷部位","th":"บริเวณยุบ","vn":"vùng lõm"},"입체감":{"en":"three-dimensional / dimension","jp":"立体感","cn":"立体感","tw":"立體感","th":"มิติ","vn":"hiệu ứng 3D"},"사전상담":{"en":"pre-consultation","jp":"事前カウンセリング","cn":"术前咨询","tw":"術前諮詢","th":"ให้คำปรึกษาก่อนทำ","vn":"tư vấn trước thực hiện"},"맞춤시술":{"en":"customized treatment","jp":"カスタマイズ施術","cn":"定制护理","tw":"客製化療程","th":"การรักษาเฉพาะบุคคล","vn":"điều trị cá nhân hóa"},"Shrink / Shurink":{"en":"Shrink / Shurink","jp":"シュリンク","cn":"修复","tw":"Shurink","th":"Shrink","vn":"Shrink"},"슈링크":{"en":"Shurink","jp":"シュリンク","cn":"舒颜提拉","tw":"Shurink","th":"Shurink","vn":"Shurink"}}}`;
let DATA = JSON.parse(RAW_DATA);

/* ===== 구글시트 런타임 데이터 연동 =====
   아래 3개 URL에 구글시트 각 탭의 "웹에 게시(CSV)" 링크를 넣으면 앱이 시트에서
   데이터를 읽어옵니다. 비어 있으면 위 내장(RAW_DATA) 데이터를 사용합니다.
   => 시트만 수정하고 페이지 새로고침하면 반영됨 (재배포 불필요). */
const SHEET_URLS = {
  tm: 'https://docs.google.com/spreadsheets/d/1vaz0HhQiuakPOiSPteFyFCqrFiLnLUQzmvI7_ck9-mQ/gviz/tq?tqx=out:csv&gid=880510946',
  glossary: 'https://docs.google.com/spreadsheets/d/1vaz0HhQiuakPOiSPteFyFCqrFiLnLUQzmvI7_ck9-mQ/gviz/tq?tqx=out:csv&gid=1301153782',
  cat: 'https://docs.google.com/spreadsheets/d/1vaz0HhQiuakPOiSPteFyFCqrFiLnLUQzmvI7_ck9-mQ/gviz/tq?tqx=out:csv&gid=2109201548',
};
const SHEET_LANGS = ['en', 'jp', 'cn', 'tw', 'th', 'vn'];
// 시트 표기(사용자 선호) → 앱 내부 코드 매핑. 신표기 jp/cn/tw/vn 우선, 구표기도 호환.
const SHEET_COL_ALIASES = { en: ['en'], jp: ['jp', 'ja'], cn: ['cn', 'zh'], tw: ['tw', 'zhtw', 'zh-tw'], th: ['th'], vn: ['vn', 'vi'] };
function csvToRows(text) {
  // RFC4180 CSV 파서 (따옴표 안의 쉼표/줄바꿈 보존)
  const rows = []; let row = []; let field = ''; let inQ = false; let i = 0;
  text = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  while (i < text.length) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQ = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === '"') { inQ = true; i++; continue; }
    if (ch === ',') { row.push(field); field = ''; i++; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
    field += ch; i++;
  }
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}
function buildFromSheets(tmRows, glossRows, catRows) {
  const hidx = (rows) => { const m = {}; (rows[0] || []).forEach((h, i) => { m[String(h).trim().toLowerCase()] = i; }); return m; };
  const colIdx = (map, lang, suffix) => { for (const code of SHEET_COL_ALIASES[lang]) { const key = (code + suffix).toLowerCase(); if (map[key] != null) return map[key]; } return null; };
  const cell = (row, i) => (i == null || row[i] == null) ? '' : String(row[i]).trim();
  const ti = hidx(tmRows); const tm = [];
  for (let r = 1; r < tmRows.length; r++) {
    const row = tmRows[r] || []; const k = cell(row, ti['k']); if (!k) continue;
    const t = {};
    for (const lang of SHEET_LANGS) {
      const e = cell(row, colIdx(ti, lang, '_e')); const d = cell(row, colIdx(ti, lang, '_d'));
      if (e) t[lang] = { e, d };
    }
    tm.push({ k, kd: cell(row, ti['kd']), c: cell(row, ti['c']), t });
  }
  const gi = hidx(glossRows); const glossary = {};
  for (let r = 1; r < glossRows.length; r++) {
    const row = glossRows[r] || []; const term = cell(row, gi['term']); if (!term) continue;
    const o = {}; for (const lang of SHEET_LANGS) { const v = cell(row, colIdx(gi, lang, '')); if (v) o[lang] = v; }
    glossary[term] = o;
  }
  const ci = hidx(catRows); const cat = {};
  for (let r = 1; r < catRows.length; r++) {
    const row = catRows[r] || []; const c = cell(row, ci['c']); if (!c) continue;
    const o = {}; for (const lang of SHEET_LANGS) { const v = cell(row, colIdx(ci, lang, '')); if (v) o[lang] = v; }
    cat[c] = o;
  }
  return { cat, tm, glossary };
}
async function loadDataFromSheets() {
  if (!SHEET_URLS.tm || !SHEET_URLS.glossary || !SHEET_URLS.cat) return null;
  try {
    const [a, b, c] = await Promise.all([
      fetch(SHEET_URLS.tm).then((r) => r.text()),
      fetch(SHEET_URLS.glossary).then((r) => r.text()),
      fetch(SHEET_URLS.cat).then((r) => r.text()),
    ]);
    const d = buildFromSheets(csvToRows(a), csvToRows(b), csvToRows(c));
    return (d.tm && d.tm.length > 0) ? d : null;
  } catch (e) {
    console.warn('구글시트 로드 실패 → 내장 데이터 사용:', e);
    return null;
  }
}


const LANGUAGES = [
  { code: 'en', native: 'English', flag: '🇺🇸', label: '영어', srcCol: 4 },
  { code: 'jp', native: '日本語', flag: '🇯🇵', label: '일본어', srcCol: 5 },
  { code: 'tw', native: '繁體中文', flag: '🇹🇼', label: '번체 (대만)', srcCol: 6 },
  { code: 'cn', native: '简体中文', flag: '🇨🇳', label: '간체 (중국)', srcCol: 7 },
  { code: 'th', native: 'ภาษาไทย', flag: '🇹🇭', label: '태국어', srcCol: 8 },
  { code: 'vn', native: 'Tiếng Việt', flag: '🇻🇳', label: '베트남어', srcCol: 9 },
];

function tokenize(text) {
  if (!text) return [];
  return text.toLowerCase().split(/[\s/()[\]{}.,，、·・~∼\-]+/).filter(t => t.length >= 2);
}

function findExactMatch(text, db) {
  const norm = (text || '').trim();
  if (!norm) return null;
  return db.tm.find(e => e.k.trim() === norm) || null;
}

function findSimilar(text, db, n = 5) {
  const tokens = new Set(tokenize(text));
  if (tokens.size === 0) return [];
  const scored = db.tm.map(e => {
    const eTokens = new Set(tokenize(e.k + ' ' + (e.kd || '')));
    let overlap = 0;
    for (const t of tokens) if (eTokens.has(t)) overlap++;
    return { entry: e, score: overlap / Math.max(tokens.size, 1) };
  }).filter(x => x.score > 0.1);
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, n).map(x => x.entry);
}

function relevantGlossary(text, db) {
  const result = {};
  for (const [ko, trans] of Object.entries(db.glossary || {})) {
    if (text && text.includes(ko)) result[ko] = trans;
  }
  return result;
}

function inferCategory(text, db) {
  for (const cat of Object.keys(db.cat || {})) if (text && text.includes(cat)) return cat;
  const lower = (text || '').toLowerCase();
  if (lower.includes('보톡스') || lower.includes('윤곽주사')) return '보톡스/윤곽주사';
  if (lower.includes('필러')) return '필러';
  if (lower.includes('실리프팅') || lower.includes('실 리프팅')) return '실리프팅';
  if (lower.includes('리프팅')) return '리프팅';
  if (lower.includes('스킨부스터') || lower.includes('스킨 부스터')) return '스킨부스터';
  if (lower.includes('레이저') || lower.includes('피코') || lower.includes('토닝')) return '피부레이저(색소/혈관/기미)';
  if (lower.includes('여드름') || lower.includes('모공') || lower.includes('흉터')) return '여드름/모공/흉터';
  if (lower.includes('제모')) return '제모';
  if (lower.includes('다이어트') || lower.includes('비만') || lower.includes('지방')) return '비만/다이어트';
  return null;
}

function formatPrice(p) {
  if (p === null || p === undefined || p === '') return null;
  const n = typeof p === 'string' ? parseInt(p.replace(/[^0-9]/g, ''), 10) : p;
  if (isNaN(n)) return null;
  return n.toLocaleString('en-US') + '원';
}

// 단건 번역 (Single mode)
async function translateSingle({ koText, koDesc, targetLangs, db }) {
  const exact = findExactMatch(koText, db);
  const similar = findSimilar(koText + ' ' + (koDesc || ''), db, 6);
  const glossary = relevantGlossary(koText + ' ' + (koDesc || ''), db);
  const category = inferCategory(koText + ' ' + (koDesc || ''), db);
  
  if (exact && targetLangs.every(l => exact.t[l])) {
    const result = {};
    for (const lang of targetLangs) {
      result[lang] = {
        event: exact.t[lang].e,
        desc: exact.t[lang].d || null,
        source: 'tm-exact',
      };
    }
    return { result, meta: { matchType: 'exact', similar: [], category } };
  }
  
  const glossaryLines = Object.entries(glossary).map(([ko, t]) => {
    const parts = targetLangs.map(l => t[l] ? `${l}: ${t[l]}` : null).filter(Boolean);
    return `  - "${ko}" → ${parts.join(', ')}`;
  }).join('\n');
  
  const categoryLine = category && db.cat[category] ? 
    `\nCATEGORY: ${category}\n` + targetLangs.map(l => `  ${l}: ${db.cat[category][l] || '?'}`).join('\n') : '';
  
  const exampleLines = similar.slice(0, 5).map((e, i) => {
    const parts = targetLangs.map(l => {
      if (!e.t[l]) return null;
      let s = `    ${l}: "${e.t[l].e}"`;
      if (e.t[l].d) s += ` | desc: "${e.t[l].d}"`;
      return s;
    }).filter(Boolean).join('\n');
    if (!parts) return null;
    let block = `Example ${i+1}:\n  KO: "${e.k}"`;
    if (e.kd) block += ` | desc: "${e.kd}"`;
    return block + '\n' + parts;
  }).filter(Boolean).join('\n\n');
  
  const systemPrompt = `You are a professional medical/cosmetic translator for 미미썸의원 (Mimisseom Clinic). Translations appear in marketing materials shown directly to international customers.

RULES:
1. Use exact glossary terms. Preserve brand names (HiTox, Coretox, Innotox, Xeomin, Jetema, Dysport, Juvelook, Rejuran, etc.).
2. Match the style of reference examples.
3. Be concise and natural for clinic marketing.
4. cn = Simplified Chinese; tw = Traditional Chinese.
5. Return ONLY valid JSON, no markdown.

GLOSSARY:
${glossaryLines || '  (none)'}
${categoryLine}

REFERENCE EXAMPLES:
${exampleLines || '(none)'}`;

  const userPrompt = `Translate to: ${targetLangs.join(', ')}.

KO EVENT: "${koText}"
${koDesc ? `KO DESC: "${koDesc}"` : ''}

JSON output:
{
${targetLangs.map(l => `  "${l}": { "event": "...", "desc": ${koDesc ? '"..."' : 'null'} }`).join(',\n')}
}`;

  const response = await fetch("/api/translate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: systemPrompt,
      messages: [{ role: "user", content: userPrompt }],
    }),
  });
  
  if (!response.ok) throw new Error(`API error: ${response.status}`);
  const data = await response.json();
  let text = data.content.map(b => b.text || '').join('').replace(/```json\s*|```\s*/g, '').trim();
  const s = text.indexOf('{'), e = text.lastIndexOf('}');
  if (s < 0 || e < 0) throw new Error('Invalid JSON');
  const parsed = JSON.parse(text.slice(s, e + 1));
  
  if (exact) {
    for (const lang of targetLangs) {
      if (exact.t[lang]) {
        parsed[lang] = { event: exact.t[lang].e, desc: exact.t[lang].d || (parsed[lang] && parsed[lang].desc) || null, source: 'tm-exact' };
      }
    }
  }
  return { result: parsed, meta: { matchType: exact ? 'partial' : 'similar', similar: similar.slice(0, 3), category, glossary } };
}

// 일괄 번역 - 한 언어, 여러 이벤트
async function translateBatchOneLang(events, lang, db) {
  // events: [{id, ko, desc, category}]
  // 글로서리 + 유사 예제 추출 (전체 batch 텍스트 기반)
  const allText = events.map(e => e.ko + ' ' + (e.desc || '')).join(' ');
  const glossary = relevantGlossary(allText, db);
  
  const glossaryLines = Object.entries(glossary).map(([ko, t]) => 
    `  - "${ko}" → ${t[lang] || '?'}`
  ).join('\n');
  
  // 각 이벤트별로 유사 예제 1-2개 모아서 batch 컨텍스트로
  const allSimilar = new Set();
  for (const ev of events) {
    const sim = findSimilar(ev.ko, db, 2);
    for (const s of sim) allSimilar.add(s);
    if (allSimilar.size >= 8) break;
  }
  
  const exampleLines = Array.from(allSimilar).slice(0, 6).map((e, i) => {
    if (!e.t[lang]) return null;
    let line = `  KO: "${e.k}"${e.kd ? ` | desc: "${e.kd}"` : ''}\n  → "${e.t[lang].e}"${e.t[lang].d ? ` | desc: "${e.t[lang].d}"` : ''}`;
    return line;
  }).filter(Boolean).join('\n\n');
  
  const langName = LANGUAGES.find(l => l.code === lang).native;
  
  const systemPrompt = `You are a professional translator for 미미썸의원 (Mimisseom Clinic), translating Korean cosmetic/medical clinic events into ${langName}. Output appears in customer-facing marketing materials.

RULES:
1. Use exact glossary terms below. Preserve brand names (HiTox, Coretox, Innotox, Xeomin, Jetema, Dysport, Juvelook, Rejuran, etc.) - do not translate them.
2. Match the style and format of reference examples.
3. Be concise, natural, and appropriate for clinic marketing.
4. ${lang === 'cn' ? 'Use Simplified Chinese characters (中国大陆).' : ''}${lang === 'tw' ? 'Use Traditional Chinese characters (繁體中文, 台灣).' : ''}
5. Return ONLY a valid JSON array. No markdown, no extra text.

GLOSSARY:
${glossaryLines || '  (none relevant)'}

REFERENCE EXAMPLES:
${exampleLines || '(none)'}`;

  const inputItems = events.map(e => ({
    id: e.id,
    ko: e.ko,
    ...(e.desc ? { desc: e.desc } : {})
  }));

  const userPrompt = `Translate each Korean event to ${langName}. Return JSON array with same ids.

INPUT:
${JSON.stringify(inputItems, null, 2)}

OUTPUT FORMAT:
[
  ${events.map(e => `{ "id": "${e.id}", "event": "...", ${e.desc ? '"desc": "..."' : '"desc": null'} }`).join(',\n  ')}
]`;

  const response = await fetch("/api/translate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: systemPrompt,
      messages: [{ role: "user", content: userPrompt }],
    }),
  });
  
  if (!response.ok) throw new Error(`API ${response.status}`);
  const data = await response.json();
  let text = data.content.map(b => b.text || '').join('').replace(/```json\s*|```\s*/g, '').trim();
  const s = text.indexOf('['), e = text.lastIndexOf(']');
  if (s < 0 || e < 0) throw new Error('Invalid JSON array');
  return JSON.parse(text.slice(s, e + 1));
}

// Excel 파싱 - 한국어 원본 파일
function parseSourceExcel(workbook) {
  const sheetName = workbook.SheetNames.find(n => n.includes('이벤트 입력')) || workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  
  // 동적으로 핵심 행 찾기 (스타일 변경에 견고)
  let titleRow = -1, langHeaderRow = -1, langSelectRow = -1, headerRow = -1;
  let branchName = '';
  
  for (let i = 0; i < Math.min(aoa.length, 20); i++) {
    const row = aoa[i] || [];
    const col0 = row[0] ? String(row[0]) : '';
    
    if (titleRow < 0 && (col0.includes('🏥') || col0.includes('외국인 HP') || col0.includes('이벤트 등록'))) {
      titleRow = i;
      // 매장명 추출: "🏥 [매장명] | ..." 또는 "🏥 광주광천점 | ..." 형식 모두 처리
      let extracted = col0.replace(/🏥/g, '').split('|')[0].trim();
      // 대괄호 안 안내 텍스트는 비워둠
      if (extracted.startsWith('[') && extracted.endsWith(']')) {
        extracted = '';
      }
      branchName = extracted;
    }
    if (langHeaderRow < 0 && col0.includes('제공 언어')) langHeaderRow = i;
    if (langSelectRow < 0 && col0.includes('서비스 언어 선택')) langSelectRow = i;
    if (headerRow < 0 && col0 === '카테고리' && row[1] && String(row[1]).includes('이벤트명')) headerRow = i;
  }
  
  // 언어 플래그 (서비스 언어 선택 행 기준) - True/False, ○, ☑/☐, ✓ 모두 인식
  const langFlags = {};
  if (langSelectRow >= 0 && aoa[langSelectRow]) {
    LANGUAGES.forEach(l => {
      const v = aoa[langSelectRow][l.srcCol];
      const sv = String(v || '').trim();
      langFlags[l.code] = (
        v === true ||
        sv === 'TRUE' || sv === 'True' || sv === 'true' ||
        sv === '○' || sv === '☑' || sv === '✓' || sv === '✅' ||
        sv === '1'
      );
    });
  } else {
    LANGUAGES.forEach(l => { langFlags[l.code] = false; });
  }
  
  // 데이터 시작 행 (헤더 다음 행)
  const dataStartRow = headerRow >= 0 ? headerRow + 1 : 5;
  
  // 데이터 행 파싱
  const events = [];
  let currentCategory = null;
  
  for (let i = dataStartRow; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const cat = row[0];
    const eventName = row[1];
    const price = row[2];
    const desc = row[3];
    
    // 카테고리 propagate
    if (cat && String(cat).trim()) {
      currentCategory = String(cat).trim();
    }
    
    // 빈 행 스킵
    if (!eventName || !String(eventName).trim()) continue;
    
    events.push({
      id: `r${i}`,
      rowIdx: i,
      categoryShownInThisRow: cat && String(cat).trim() ? String(cat).trim() : null,
      category: currentCategory,
      ko: String(eventName).trim(),
      price: price !== null && price !== undefined && price !== '' ? price : null,
      desc: desc && String(desc).trim() ? String(desc).trim() : null,
    });
  }
  
  return { 
    branchName, 
    langFlags, 
    events, 
    sheetName, 
    totalRows: aoa.length,
    titleRow,
    langHeaderRow,
    langSelectRow,
    headerRow,
    dataStartRow,
  };
}

// 출력 Excel 생성 - 한 언어
function buildOutputWorkbook(sourceArrayBuffer, sourceParsed, translations, lang, db) {
  // 새 워크북 생성 (원본 바이트에서 fresh read - 스타일 보존)
  const wb = XLSX.read(sourceArrayBuffer, { type: 'array', cellStyles: true });
  const ws = wb.Sheets[sourceParsed.sheetName];
  
  const catMap = db.cat || {};
  
  // 기본 alignment - 셀에 wrap_text 적용 (긴 텍스트 표시용)
  const wrapAlign = { vertical: 'center', wrapText: true };
  const wrapAlignLeft = { vertical: 'center', wrapText: true, horizontal: 'left', indent: 1 };
  
  for (const event of sourceParsed.events) {
    const trans = translations[event.id];
    if (!trans) continue;
    
    const rowIdx = event.rowIdx;
    
    // 카테고리 (셀이 채워진 행만)
    if (event.categoryShownInThisRow) {
      const catTrans = catMap[event.category];
      const bilingualCat = catTrans && catTrans[lang] 
        ? `${event.category}\n${catTrans[lang]}` 
        : event.category;
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
      const existing = ws[ref] || {};
      // 스타일이 있으면 wrap_text 보장, 없으면 기본 추가
      const newStyle = existing.s ? 
        { ...existing.s, alignment: { ...(existing.s.alignment || {}), ...wrapAlignLeft } } :
        { alignment: wrapAlignLeft };
      ws[ref] = { ...existing, v: bilingualCat, t: 's', w: bilingualCat, s: newStyle };
    }
    
    // 이벤트명 (col 1)
    const eventRef = XLSX.utils.encode_cell({ r: rowIdx, c: 1 });
    const existingEvent = ws[eventRef] || {};
    const eventStyle = existingEvent.s ?
      { ...existingEvent.s, alignment: { ...(existingEvent.s.alignment || {}), ...wrapAlignLeft } } :
      { alignment: wrapAlignLeft };
    ws[eventRef] = { ...existingEvent, v: trans.event || '', t: 's', w: trans.event || '', s: eventStyle };
    
    // 가격 (col 2) - 포맷된 문자열로
    if (event.price !== null) {
      const priceFormatted = formatPrice(event.price);
      if (priceFormatted) {
        const priceRef = XLSX.utils.encode_cell({ r: rowIdx, c: 2 });
        const existingPrice = ws[priceRef] || {};
        const priceStyle = existingPrice.s ?
          { ...existingPrice.s, alignment: { vertical: 'center', horizontal: 'right', wrapText: true, indent: 1 }, numFmt: '@' } :
          { alignment: { vertical: 'center', horizontal: 'right', wrapText: true, indent: 1 } };
        ws[priceRef] = { ...existingPrice, v: priceFormatted, t: 's', w: priceFormatted, s: priceStyle };
      }
    }
    
    // 설명 (col 3)
    if (trans.desc) {
      const descRef = XLSX.utils.encode_cell({ r: rowIdx, c: 3 });
      const existingDesc = ws[descRef] || {};
      const descStyle = existingDesc.s ?
        { ...existingDesc.s, alignment: { ...(existingDesc.s.alignment || {}), ...wrapAlignLeft } } :
        { alignment: wrapAlignLeft };
      ws[descRef] = { ...existingDesc, v: trans.desc, t: 's', w: trans.desc, s: descStyle };
    }
  }
  
  // 카테고리 기준표 시트는 건드리지 않음 - 본문에 이미 이중언어 카테고리 표기 있음
  
  return wb;
}

function downloadWorkbook(wb, filename) {
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ── 의존성 없는 ZIP(STORE) 생성: 여러 xlsx를 하나의 zip으로 묶기 ──
function _crc32(bytes) {
  let table = _crc32._t;
  if (!table) {
    table = _crc32._t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
  }
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xFF];
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function _concat(parts) {
  let total = 0; for (const p of parts) total += p.length;
  const out = new Uint8Array(total); let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
function createZipBlob(files) {
  const enc = new TextEncoder();
  const u16 = (n) => new Uint8Array([n & 0xFF, (n >>> 8) & 0xFF]);
  const u32 = (n) => new Uint8Array([n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]);
  const now = new Date();
  const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  const localParts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const nameBytes = enc.encode(f.name);
    const data = f.data;
    const crc = _crc32(data);
    const size = data.length;
    const lh = _concat([
      u32(0x04034b50), u16(20), u16(0x0800), u16(0),
      u16(dosTime), u16(dosDate), u32(crc), u32(size), u32(size),
      u16(nameBytes.length), u16(0), nameBytes,
    ]);
    localParts.push(lh, data);
    central.push(_concat([
      u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0),
      u16(dosTime), u16(dosDate), u32(crc), u32(size), u32(size),
      u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0),
      u32(0), u32(offset), nameBytes,
    ]));
    offset += lh.length + data.length;
  }
  const centralBytes = _concat(central);
  const end = _concat([
    u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(centralBytes.length), u32(offset), u16(0),
  ]);
  return new Blob([...localParts, centralBytes, end], { type: 'application/zip' });
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getOutputFilename(srcFilename, lang) {
  const langSuffix = {
    'en': '_영어',
    'jp': '_일본어',
    'tw': '_대만어',
    'cn': '_중국어',
    'th': '_태국어',
    'vn': '_베트남어',
  }[lang] || `_${lang}`;
  
  const base = srcFilename.replace(/\.xlsx?$/i, '');
  return `${base}${langSuffix}.xlsx`;
}

// 샘플 양식 (디자인된 xlsx 파일을 base64로 임베드)
const SAMPLE_TEMPLATE_B64 = "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";

function base64ToBlob(b64, mime) {
  const byteChars = atob(b64);
  const len = byteChars.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = byteChars.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

function downloadSampleTemplate() {
  const blob = base64ToBlob(SAMPLE_TEMPLATE_B64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '_외국인이벤트_샘플양식.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ========= UI Components =========

function Header({ tab, setTab }) {
  return (
    <header className="border-b" style={{ borderColor: '#E8DFD0' }}>
      <div className="max-w-6xl mx-auto px-6 py-5">
        <div className="flex items-center justify-between mb-4">
          <div>
            <h1 className="font-display text-2xl md:text-3xl font-medium tracking-tight" style={{ color: '#3D2817' }}>
              미미썸의원 <span className="italic font-normal" style={{ color: '#B8804E' }}>Global</span>
            </h1>
            <p className="text-xs mt-1 tracking-wide" style={{ color: '#86715A' }}>
              EVENT TRANSLATOR · 외국인 이벤트 번역 도구
            </p>
          </div>
        </div>
        <nav className="flex gap-1 -mb-px">
          {[
            { id: 'batch', label: '엑셀 일괄 번역', sub: 'Bulk' },
            { id: 'single', label: '단건 번역', sub: 'Single' },
          ].map(t => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className="px-4 py-2.5 text-sm transition-colors relative"
              style={{
                color: tab === t.id ? '#3D2817' : '#86715A',
                borderBottom: tab === t.id ? '2px solid #B8804E' : '2px solid transparent',
                fontWeight: tab === t.id ? 600 : 400,
              }}
            >
              {t.label}
              <span className="ml-1.5 text-xs opacity-50 italic" style={{ fontFamily: "'Fraunces', serif" }}>{t.sub}</span>
            </button>
          ))}
        </nav>
      </div>
    </header>
  );
}

function BatchMode() {
  const fileInputRef = React.useRef(null);
  const [file, setFile] = React.useState(null);
  const [parsed, setParsed] = React.useState(null);
  const [parseError, setParseError] = React.useState(null);
  const [selectedLangs, setSelectedLangs] = React.useState({
    en: true, jp: false, tw: false, cn: false, th: false, vn: true,
  });
  const [translating, setTranslating] = React.useState(false);
  const [progress, setProgress] = React.useState(null);
  const [outputs, setOutputs] = React.useState({});
  const [translationError, setTranslationError] = React.useState(null);
  const [isDragging, setIsDragging] = React.useState(false);

  const handleFile = async (f) => {
    if (!f) return;
    setFile(f);
    setParseError(null);
    setOutputs({});
    setTranslationError(null);
    
    try {
      const arrayBuffer = await f.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
      const data = parseSourceExcel(wb);
      data.workbook = wb;
      data.arrayBuffer = arrayBuffer;
      data.filename = f.name;
      setParsed(data);
      
      // 원본 파일의 언어 선택 자동 적용
      const newSel = { ...selectedLangs };
      let hasAnyTrue = false;
      LANGUAGES.forEach(l => {
        if (data.langFlags[l.code]) hasAnyTrue = true;
      });
      if (hasAnyTrue) {
        LANGUAGES.forEach(l => { newSel[l.code] = !!data.langFlags[l.code]; });
        setSelectedLangs(newSel);
      }
    } catch (e) {
      console.error(e);
      setParseError('엑셀 파싱 실패: ' + e.message);
      setParsed(null);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  };

  const targetLangs = LANGUAGES.filter(l => selectedLangs[l.code]).map(l => l.code);

  const eventsByCategory = React.useMemo(() => {
    if (!parsed) return {};
    const grouped = {};
    for (const ev of parsed.events) {
      const cat = ev.category || '(미분류)';
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(ev);
    }
    return grouped;
  }, [parsed]);

  const tmCoverage = React.useMemo(() => {
    if (!parsed) return null;
    const stats = { total: parsed.events.length, perLang: {} };
    targetLangs.forEach(lang => {
      let hits = 0;
      for (const ev of parsed.events) {
        const tm = findExactMatch(ev.ko, DATA);
        if (tm && tm.t[lang]) hits++;
      }
      stats.perLang[lang] = hits;
    });
    return stats;
  }, [parsed, targetLangs.join(',')]);

  const handleTranslate = async () => {
    if (!parsed || targetLangs.length === 0) return;
    setTranslating(true);
    setOutputs({});
    setTranslationError(null);
    
    try {
      // 언어별 번역 결과 캐시
      const allTranslations = {}; // {lang: {eventId: {event, desc}}}
      for (const lang of targetLangs) allTranslations[lang] = {};
      
      // 1. TM 매치 먼저
      const needsApi = {}; // {lang: [event]}
      for (const lang of targetLangs) needsApi[lang] = [];
      
      for (const ev of parsed.events) {
        const tm = findExactMatch(ev.ko, DATA);
        for (const lang of targetLangs) {
          if (tm && tm.t[lang]) {
            allTranslations[lang][ev.id] = {
              event: tm.t[lang].e,
              desc: tm.t[lang].d || (ev.desc ? null : null), // TM의 desc 우선, 없고 ev.desc도 없으면 null
              source: 'tm',
            };
            // ev.desc는 있는데 tm.t[lang].d가 없는 경우 → API로 desc만 번역해야 함
            if (ev.desc && !tm.t[lang].d) {
              // partial TM hit: event matched, desc needs API
              needsApi[lang].push({ ...ev, _partialTM: true });
            }
          } else {
            needsApi[lang].push(ev);
          }
        }
      }
      
      // 2. API 호출 (언어별, batch)
      const batchSize = 4;
      let totalApiCalls = 0;
      let doneApiCalls = 0;
      for (const lang of targetLangs) {
        totalApiCalls += Math.ceil(needsApi[lang].length / batchSize);
      }
      
      setProgress({ phase: 'translating', total: totalApiCalls, done: 0, currentLang: null });
      
      for (const lang of targetLangs) {
        const eventsForLang = needsApi[lang];
        for (let i = 0; i < eventsForLang.length; i += batchSize) {
          const batch = eventsForLang.slice(i, i + batchSize);
          setProgress({ phase: 'translating', total: totalApiCalls, done: doneApiCalls, currentLang: lang });
          
          try {
            const results = await translateBatchOneLang(
              batch.map(e => ({ id: e.id, ko: e.ko, desc: e.desc })),
              lang,
              DATA
            );
            
            for (const r of results) {
              if (!allTranslations[lang][r.id]) {
                allTranslations[lang][r.id] = { event: r.event, desc: r.desc, source: 'api' };
              } else {
                // 이미 TM에서 event 매치된 경우: event는 TM 값 유지, desc만 업데이트
                if (r.desc) allTranslations[lang][r.id].desc = r.desc;
              }
            }
          } catch (e) {
            console.error('Batch failed:', e);
            // 실패한 배치는 건너뛰고 계속 진행 (개별 이벤트는 빈 상태)
            for (const ev of batch) {
              if (!allTranslations[lang][ev.id]) {
                allTranslations[lang][ev.id] = { event: '[번역 실패] ' + ev.ko, desc: ev.desc, source: 'failed' };
              }
            }
          }
          
          doneApiCalls++;
          setProgress({ phase: 'translating', total: totalApiCalls, done: doneApiCalls, currentLang: lang });
        }
      }
      
      // 3. 출력 워크북 생성
      setProgress({ phase: 'building', total: targetLangs.length, done: 0 });
      const newOutputs = {};
      for (let i = 0; i < targetLangs.length; i++) {
        const lang = targetLangs[i];
        const wb = buildOutputWorkbook(parsed.arrayBuffer, parsed, allTranslations[lang], lang, DATA);
        const filename = getOutputFilename(parsed.filename, lang);
        newOutputs[lang] = { wb, filename };
        setProgress({ phase: 'building', total: targetLangs.length, done: i + 1 });
      }
      
      setOutputs(newOutputs);
      setProgress(null);
    } catch (e) {
      console.error(e);
      setTranslationError(e.message || String(e));
      setProgress(null);
    } finally {
      setTranslating(false);
    }
  };

  const handleDownload = (lang) => {
    const out = outputs[lang];
    if (out) downloadWorkbook(out.wb, out.filename);
  };
  
  const handleDownloadAll = () => {
    const entries = Object.values(outputs);
    if (entries.length === 0) return;
    if (entries.length === 1) {
      downloadWorkbook(entries[0].wb, entries[0].filename);
      return;
    }
    // 2건 이상: 하나의 zip으로 묶어서 다운로드
    const files = entries.map((out) => ({
      name: out.filename,
      data: new Uint8Array(XLSX.write(out.wb, { type: 'array', bookType: 'xlsx' })),
    }));
    const base = (parsed && parsed.filename ? parsed.filename : '번역결과').replace(/\.xlsx?$/i, '');
    downloadBlob(createZipBlob(files), `${base}_번역.zip`);
  };

  return (
    <div className="max-w-6xl mx-auto px-6 py-8">
      {/* 1. File upload */}
      {!parsed && (
        <div>
          <div
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className="cursor-pointer rounded-lg p-12 text-center transition-all"
            style={{
              border: '2px dashed ' + (isDragging ? '#B8804E' : '#D4C9B5'),
              backgroundColor: isDragging ? '#FAF6EE' : '#FFFEFC',
            }}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => handleFile(e.target.files[0])}
            />
            <div className="text-5xl mb-4 opacity-30">📁</div>
            <p className="font-display text-xl mb-2" style={{ color: '#3D2817' }}>
              한국어 이벤트 엑셀 파일을 올려주세요
            </p>
            <p className="text-sm" style={{ color: '#86715A' }}>
              드래그해서 놓거나 클릭해서 선택 · .xlsx 파일
            </p>
            <p className="text-xs mt-4" style={{ color: '#B8804E' }}>
              업로드한 형식 그대로, 선택한 각 언어별 엑셀 파일로 내려드립니다
            </p>
          </div>

          {/* 샘플 양식 다운로드 */}
          <div className="mt-4 p-4 rounded-lg flex items-center justify-between gap-4" 
            style={{ backgroundColor: '#FAF6EE', border: '1px solid #E8DFD0' }}>
            <div className="flex-1">
              <p className="text-sm font-medium mb-0.5" style={{ color: '#3D2817' }}>
                처음이세요? 샘플 양식 받으세요
              </p>
              <p className="text-xs" style={{ color: '#86715A' }}>
                정해진 형식대로 입력해야 오류 없이 번역돼요. 예시 12개 포함된 양식을 받아서 채워보세요.
              </p>
            </div>
            <button
              onClick={(e) => { e.stopPropagation(); downloadSampleTemplate(); }}
              className="px-4 py-2 text-xs rounded-md transition whitespace-nowrap"
              style={{ backgroundColor: '#B8804E', color: '#FFFEFC' }}
            >
              📥 샘플 양식 다운로드
            </button>
          </div>
        </div>
      )}
      
      {parseError && (
        <div className="mt-4 p-3 rounded-md text-xs" style={{ backgroundColor: '#FAE5E5', color: '#8B3A3A' }}>
          {parseError}
        </div>
      )}

      {/* 2. Preview */}
      {parsed && (
        <div className="space-y-6">
          <div className="p-5 rounded-lg" style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}>
            <div className="flex items-start justify-between mb-3">
              <div>
                <div className="text-xs tracking-wide uppercase mb-1" style={{ color: '#86715A' }}>업로드된 파일</div>
                <div className="font-medium text-base" style={{ color: '#3D2817' }}>{parsed.filename}</div>
                <div className="text-xs mt-1" style={{ color: '#86715A' }}>
                  {parsed.branchName && <span>매장: <strong>{parsed.branchName}</strong> · </span>}
                  이벤트 <strong>{parsed.events.length}개</strong> · 
                  카테고리 <strong>{Object.keys(eventsByCategory).length}종</strong>
                </div>
              </div>
              <button
                onClick={() => { setParsed(null); setFile(null); setOutputs({}); }}
                className="text-xs px-3 py-1 rounded transition"
                style={{ color: '#86715A', border: '1px solid #E8DFD0' }}
              >
                다른 파일
              </button>
            </div>
            
            <details className="mt-3">
              <summary className="text-xs cursor-pointer" style={{ color: '#B8804E' }}>
                카테고리별 이벤트 보기 ▾
              </summary>
              <div className="mt-3 space-y-2 max-h-72 overflow-y-auto scrollbar-thin">
                {Object.entries(eventsByCategory).map(([cat, evs]) => (
                  <div key={cat} className="text-xs">
                    <div className="font-medium pb-1" style={{ color: '#3D2817' }}>
                      {cat} <span style={{ color: '#86715A' }}>({evs.length})</span>
                    </div>
                    <div className="pl-3 space-y-0.5" style={{ color: '#5C4A38' }}>
                      {evs.slice(0, 5).map(ev => <div key={ev.id}>· {ev.ko}{ev.price ? ` (${formatPrice(ev.price)})` : ''}</div>)}
                      {evs.length > 5 && <div style={{ color: '#86715A' }}>… 외 {evs.length - 5}개</div>}
                    </div>
                  </div>
                ))}
              </div>
            </details>
          </div>

          {/* 3. Language selection */}
          <div className="p-5 rounded-lg" style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}>
            <div className="flex items-center justify-between mb-4">
              <h2 className="font-display text-lg font-medium" style={{ color: '#3D2817' }}>
                <span style={{ color: '#B8804E' }}>→</span> 번역할 언어 선택
              </h2>
              <span className="text-xs italic" style={{ color: '#86715A' }}>
                {Object.values(parsed.langFlags).some(Boolean) ? '원본 파일의 선택을 적용했어요' : '원하는 언어를 선택해주세요'}
              </span>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2 mb-3">
              {LANGUAGES.map(lang => {
                const isSelected = selectedLangs[lang.code];
                const tmHit = tmCoverage && tmCoverage.perLang[lang.code];
                return (
                  <label key={lang.code}
                    className="checkbox-tile cursor-pointer flex flex-col items-center gap-1 py-3 px-2 rounded-md text-center"
                    style={{
                      backgroundColor: isSelected ? '#3D2817' : '#FAF6EE',
                      color: isSelected ? '#F7F3EC' : '#3D2817',
                      border: '1px solid ' + (isSelected ? '#3D2817' : '#E8DFD0'),
                    }}
                  >
                    <input type="checkbox" checked={isSelected}
                      onChange={() => setSelectedLangs(s => ({ ...s, [lang.code]: !s[lang.code] }))}
                      className="sr-only" />
                    <span className="text-2xl">{lang.flag}</span>
                    <span className="text-xs font-medium">{lang.label}</span>
                    {isSelected && tmHit !== undefined && (
                      <span className="text-[10px] opacity-75">
                        TM {tmHit}/{tmCoverage.total}
                      </span>
                    )}
                  </label>
                );
              })}
            </div>
            {tmCoverage && targetLangs.length > 0 && (
              <p className="text-xs mt-2" style={{ color: '#86715A' }}>
                <span style={{ color: '#5C8F5C' }}>●</span> TM 매치 = 과거 번역 그대로 사용 (API 호출 없음, 빠르고 정확). 
                나머지만 AI 번역됩니다.
              </p>
            )}
          </div>

          {/* 4. Action */}
          {Object.keys(outputs).length === 0 && (
            <div className="flex items-center gap-3">
              <button
                onClick={handleTranslate}
                disabled={translating || targetLangs.length === 0}
                className="px-6 py-3 text-sm font-medium tracking-wide rounded-md transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                style={{ backgroundColor: '#3D2817', color: '#F7F3EC' }}
              >
                {translating ? '번역 중…' : `${targetLangs.length}개 언어로 번역 시작 →`}
              </button>
              {translationError && (
                <span className="text-xs" style={{ color: '#B85E5E' }}>오류: {translationError}</span>
              )}
            </div>
          )}

          {/* 5. Progress */}
          {progress && (
            <div className="p-4 rounded-lg" style={{ backgroundColor: '#FAF6EE', border: '1px solid #E8DFD0' }}>
              <div className="flex items-center justify-between text-sm mb-2">
                <span style={{ color: '#3D2817' }}>
                  {progress.phase === 'translating' ? '번역 중' : '엑셀 파일 생성 중'}
                  {progress.currentLang && ` · ${LANGUAGES.find(l => l.code === progress.currentLang)?.native}`}
                </span>
                <span className="text-xs" style={{ color: '#86715A' }}>
                  {progress.done} / {progress.total}
                </span>
              </div>
              <div className="h-2 rounded-full overflow-hidden" style={{ backgroundColor: '#E8DFD0' }}>
                <div className="h-full transition-all"
                  style={{
                    backgroundColor: '#B8804E',
                    width: `${Math.round((progress.done / Math.max(progress.total, 1)) * 100)}%`,
                  }}
                />
              </div>
            </div>
          )}

          {/* 6. Outputs */}
          {Object.keys(outputs).length > 0 && (
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <h2 className="font-display text-lg font-medium" style={{ color: '#3D2817' }}>
                  ✓ 번역 완료
                </h2>
                <button
                  onClick={handleDownloadAll}
                  className="px-4 py-2 text-xs rounded-md transition"
                  style={{ backgroundColor: '#3D2817', color: '#F7F3EC' }}
                >
                  전체 다운로드
                </button>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                {Object.entries(outputs).map(([lang, out]) => {
                  const langInfo = LANGUAGES.find(l => l.code === lang);
                  return (
                    <button
                      key={lang}
                      onClick={() => handleDownload(lang)}
                      className="flex items-center justify-between p-4 rounded-md transition text-left hover:shadow-sm"
                      style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}
                    >
                      <div className="flex items-center gap-3">
                        <span className="text-2xl">{langInfo.flag}</span>
                        <div>
                          <div className="text-sm font-medium" style={{ color: '#3D2817' }}>{langInfo.native}</div>
                          <div className="text-[11px]" style={{ color: '#86715A' }}>{out.filename}</div>
                        </div>
                      </div>
                      <span className="text-xs px-3 py-1.5 rounded" style={{ backgroundColor: '#F0E5D2', color: '#5C4A38' }}>
                        다운로드 ↓
                      </span>
                    </button>
                  );
                })}
              </div>
              <button
                onClick={() => { setOutputs({}); setTranslationError(null); }}
                className="text-xs mt-2"
                style={{ color: '#86715A' }}
              >
                ← 다시 번역하기
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function SingleMode() {
  const [koText, setKoText] = React.useState('');
  const [koDesc, setKoDesc] = React.useState('');
  const [price, setPrice] = React.useState('');
  const [selectedLangs, setSelectedLangs] = React.useState({
    en: true, jp: true, tw: true, cn: true, th: true, vn: true,
  });
  const [results, setResults] = React.useState(null);
  const [meta, setMeta] = React.useState(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState(null);
  const [copiedKey, setCopiedKey] = React.useState(null);
  
  const targetLangs = LANGUAGES.filter(l => selectedLangs[l.code]).map(l => l.code);

  const handleTranslate = async () => {
    if (!koText.trim()) { setError('한국어를 입력해주세요.'); return; }
    if (targetLangs.length === 0) { setError('언어를 선택해주세요.'); return; }
    setError(null); setIsLoading(true); setResults(null);
    try {
      const { result, meta } = await translateSingle({
        koText: koText.trim(), koDesc: koDesc.trim() || null,
        targetLangs, db: DATA,
      });
      setResults(result); setMeta(meta);
    } catch (e) { setError(e.message); }
    finally { setIsLoading(false); }
  };

  const copyToClipboard = async (text, key) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(null), 1500);
    } catch (e) { console.error(e); }
  };

  return (
    <div className="max-w-6xl mx-auto px-6 py-8">
      <section className="mb-6 p-6 rounded-lg" style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}>
        <div className="space-y-4">
          <div>
            <label className="block text-xs mb-1.5 tracking-wide uppercase" style={{ color: '#86715A' }}>이벤트명 (필수)</label>
            <textarea value={koText} onChange={(e) => setKoText(e.target.value)} rows={2}
              placeholder="예: 표정주름보톡스 1부위 (국산 하이톡스)"
              className="w-full px-4 py-3 text-base rounded-md outline-none resize-none"
              style={{ border: '1px solid #E8DFD0', backgroundColor: '#FAF6EE', fontFamily: "'Noto Sans KR', sans-serif" }} />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="md:col-span-2">
              <label className="block text-xs mb-1.5 tracking-wide uppercase" style={{ color: '#86715A' }}>설명 (선택)</label>
              <input type="text" value={koDesc} onChange={(e) => setKoDesc(e.target.value)}
                placeholder="이마/미간/눈가/콧대/콧등/자갈턱 중 택1"
                className="w-full px-4 py-2.5 text-sm rounded-md outline-none"
                style={{ border: '1px solid #E8DFD0', backgroundColor: '#FAF6EE', fontFamily: "'Noto Sans KR', sans-serif" }} />
            </div>
            <div>
              <label className="block text-xs mb-1.5 tracking-wide uppercase" style={{ color: '#86715A' }}>가격 (선택)</label>
              <input type="text" value={price} onChange={(e) => setPrice(e.target.value.replace(/[^0-9]/g, ''))}
                placeholder="15000"
                className="w-full px-4 py-2.5 text-sm rounded-md outline-none"
                style={{ border: '1px solid #E8DFD0', backgroundColor: '#FAF6EE' }} />
            </div>
          </div>
        </div>
      </section>
      
      <section className="mb-6 p-6 rounded-lg" style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}>
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
          {LANGUAGES.map(lang => (
            <label key={lang.code}
              className="checkbox-tile cursor-pointer flex flex-col items-center gap-1 py-3 px-2 rounded-md text-center"
              style={{
                backgroundColor: selectedLangs[lang.code] ? '#3D2817' : '#FAF6EE',
                color: selectedLangs[lang.code] ? '#F7F3EC' : '#3D2817',
                border: '1px solid ' + (selectedLangs[lang.code] ? '#3D2817' : '#E8DFD0'),
              }}>
              <input type="checkbox" checked={selectedLangs[lang.code]}
                onChange={() => setSelectedLangs(s => ({ ...s, [lang.code]: !s[lang.code] }))}
                className="sr-only" />
              <span className="text-2xl">{lang.flag}</span>
              <span className="text-xs font-medium">{lang.label}</span>
            </label>
          ))}
        </div>
      </section>

      <div className="mb-6 flex items-center gap-3">
        <button onClick={handleTranslate}
          disabled={isLoading || !koText.trim() || targetLangs.length === 0}
          className="px-6 py-3 text-sm font-medium rounded-md transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          style={{ backgroundColor: '#3D2817', color: '#F7F3EC' }}>
          {isLoading ? '번역 중…' : '번역 →'}
        </button>
        {error && <span className="text-xs" style={{ color: '#B85E5E' }}>{error}</span>}
      </div>

      {meta && meta.matchType === 'exact' && (
        <div className="mb-4 p-3 rounded-md text-xs" style={{ backgroundColor: '#F0E5D2', color: '#5C4A38' }}>
          ✓ 과거 승인된 번역 사용 (번역 메모리)
        </div>
      )}

      {isLoading && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {targetLangs.map(l => <div key={l} className="h-32 rounded-md shimmer"></div>)}
        </div>
      )}

      {results && !isLoading && (
        <section className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {targetLangs.map((langCode, idx) => {
            const lang = LANGUAGES.find(l => l.code === langCode);
            const r = results[langCode];
            if (!r) return null;
            const fontClass = langCode === 'jp' ? 'font-ja' : langCode === 'cn' ? 'font-zh' : langCode === 'tw' ? 'font-zh-tw' : langCode === 'th' ? 'font-th' : '';
            return (
              <div key={langCode} className="p-5 rounded-lg fade-in-up"
                style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0', animationDelay: (idx * 60) + 'ms' }}>
                <div className="flex items-center justify-between mb-3 pb-3" style={{ borderBottom: '1px solid #F0E5D2' }}>
                  <div className="flex items-center gap-2">
                    <span className="text-lg">{lang.flag}</span>
                    <span className="text-xs tracking-wide uppercase" style={{ color: '#86715A' }}>{lang.native}</span>
                    {r.source === 'tm-exact' && <span className="text-[10px] px-1.5 py-0.5 rounded" style={{ backgroundColor: '#E8E1D5', color: '#5C4A38' }}>TM</span>}
                  </div>
                  <button onClick={() => copyToClipboard(r.event + (r.desc ? '\n' + r.desc : ''), langCode)}
                    className="text-xs" style={{ color: copiedKey === langCode ? '#5C8F5C' : '#B8804E' }}>
                    {copiedKey === langCode ? '✓ 복사됨' : '복사'}
                  </button>
                </div>
                <div className={fontClass}>
                  <p className="text-base font-medium leading-snug mb-2" style={{ color: '#1a1614' }}>{r.event}</p>
                  {r.desc && <p className="text-sm" style={{ color: '#5C4A38' }}>{r.desc}</p>}
                  {price && <p className="text-sm font-medium mt-2 pt-2" style={{ color: '#3D2817', borderTop: '1px solid #F0E5D2' }}>{formatPrice(price)}</p>}
                </div>
              </div>
            );
          })}
        </section>
      )}
    </div>
  );
}

function App() {
  const [tab, setTab] = React.useState('batch');
  const [sheetSynced, setSheetSynced] = React.useState(false);
  React.useEffect(() => {
    loadDataFromSheets().then((d) => { if (d) { DATA = d; setSheetSynced(true); } });
  }, []);

  return (
    <div className="min-h-screen" style={{ 
      backgroundColor: '#F7F3EC', 
      fontFamily: "'Manrope', 'Noto Sans KR', sans-serif",
      color: '#1a1614',
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Manrope:wght@300;400;500;600;700&family=Noto+Sans+KR:wght@300;400;500;700&family=Noto+Sans+JP:wght@300;400;500&family=Noto+Sans+SC:wght@300;400;500&family=Noto+Sans+TC:wght@300;400;500&family=Noto+Sans+Thai:wght@300;400;500&display=swap');
        .font-display { font-family: 'Fraunces', 'Noto Sans KR', serif; }
        .font-zh { font-family: 'Noto Sans SC', sans-serif; }
        .font-zh-tw { font-family: 'Noto Sans TC', sans-serif; }
        .font-ja { font-family: 'Noto Sans JP', sans-serif; }
        .font-th { font-family: 'Noto Sans Thai', sans-serif; }
        .checkbox-tile { transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1); }
        .checkbox-tile:hover { transform: translateY(-1px); }
        @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
        .shimmer { background: linear-gradient(90deg, #E8E1D5 25%, #F5F0E6 50%, #E8E1D5 75%); background-size: 200% 100%; animation: shimmer 1.6s infinite; }
        @keyframes fade-in-up { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .fade-in-up { animation: fade-in-up 0.4s ease-out forwards; }
        .scrollbar-thin::-webkit-scrollbar { width: 6px; }
        .scrollbar-thin::-webkit-scrollbar-track { background: transparent; }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: #D4C9B5; border-radius: 3px; }
      `}</style>
      
      <Header tab={tab} setTab={setTab} />
      
      <main>
        {tab === 'batch' && <BatchMode />}
        {tab === 'single' && <SingleMode />}
      </main>
      
      <footer className="border-t mt-12 py-6" style={{ borderColor: '#E8DFD0' }}>
        <div className="max-w-6xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-2 text-xs" style={{ color: '#86715A' }}>
          <span>번역 메모리 {DATA.tm.length}건 · 용어집 {Object.keys(DATA.glossary).length}건 보유</span>
          <span style={{ fontFamily: "'Fraunces', serif" }} className="italic">미미썸의원 · 외국인 손님을 위한 정확한 번역</span>
        </div>
      </footer>
    </div>
  );
}

export default App;

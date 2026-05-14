import React from 'react';
import * as XLSX from 'xlsx';

const RAW_DATA = `{"cat":{"보톡스/윤곽주사":{"en":"Botox / Face Contouring Injection","ja":"ボトックス・輪郭注射","zh":"肉毒素 / 轮廓针","zh-TW":"肉毒素 / 輪廓針","th":"โบท็อกซ์ / ฉีดแฟต","vi":"Botox / Tiêm thon gọn line khuôn mặt"},"필러":{"en":"Filler","ja":"フィラー","zh":"填充剂","zh-TW":"填充劑","th":"ฟิลเลอร์","vi":"Filler"},"리프팅":{"en":"Lifting","ja":"リフトアップ","zh":"提升","zh-TW":"提升","th":"ยกกระชับ","vi":"Nâng cơ"},"스킨부스터":{"en":"Skin Booster","ja":"スキンブースター","zh":"皮肤增强剂","zh-TW":"皮膚增強劑","th":"สกินบูสเตอร์","vi":"Skin Booster"},"줄기세포":{"en":"Stem Cells","ja":"幹細胞","zh":"干细胞","zh-TW":"幹細胞","th":"สเต็มเซลล์","vi":"Tế bào gốc"},"실리프팅":{"en":"Thread Lifting","ja":"糸リフト","zh":"线雕","zh-TW":"線雕","th":"ร้อยไหม","vi":"Căng chỉ"},"피부레이저(색소/혈관/기미)":{"en":"Skin Laser Treatment (Pigmentation / Vascular / Melasma)","ja":"皮膚レーザー（色素・血管・シミ）","zh":"皮肤激光 (色素 / 血管 / 黄褐斑)","zh-TW":"皮膚雷射 (色素 / 血管 / 黃褐斑)","th":"เลเซอร์ผิว (เม็ดสี / รอยแดง / ฝ้า)","vi":"Laser da (sắc tố / mạch máu / nám)"},"여드름/모공/흉터":{"en":"Acne / Pores / Scars","ja":"ニキビ・毛穴・傷跡","zh":"痤疮 / 毛孔 / 疤痕","zh-TW":"痤瘡 / 毛孔 / 疤痕","th":"สิว / รูขุมขน / หลุมสิว","vi":"Mụn trứng cá / Lỗ chân lông / Sẹo"},"제모":{"en":"Hair Removal","ja":"脱毛","zh":"脱毛","zh-TW":"脫毛","th":"กำจัดขน","vi":"Triệt lông"},"비만/다이어트":{"en":"Weight Loss / Obesity Clinic","ja":"肥満・ダイエット","zh":"肥胖 / 减肥","zh-TW":"肥胖 / 減肥","th":"ลดน้ำหนัก / รูปร่าง","vi":"Béo phì / Giảm cân"},"성형수술":{"en":"Cosmetic Surgery","ja":"美容整形手術","zh":"整形手术","zh-TW":"整形手術","th":"ศัลยกรรม","vi":"Phẫu thuật thẩm mỹ"},"스킨케어":{"en":"Skin Care","ja":"スキンケア","zh":"皮肤护理","zh-TW":"皮膚護理","th":"ดูแลผิว","vi":"Chăm sóc da"},"수액":{"en":"IV Drip","ja":"点滴","zh":"营养点滴","zh-TW":"營養點滴","th":"วิตามินดริป","vi":"Truyền dịch dinh dưỡng"}},"tm":[{"k":"표정주름보톡스 1부위 (국산 하이톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) – Hutox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"ja":{"e":"表情じわボトックスの1部位（韓国製のハイトックス）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 HiTox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Hitox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Hyttox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (국산 코어톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area)  Coretox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"ja":{"e":"表情じわボトックスの1部位（韓国製のコアトックス）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Coretox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Coretox)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Coretox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (국산 이노톡스)","kd":"세계 유일 액상 보톡스, 이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Innotox (Korean)","d":"World's only liquid Botox,  Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"ja":{"e":"表情じわボトックスの1部位（韓国製のイノトックス）","d":"世界唯一の液状型ボトックス おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Innotox）","d":"世界唯一液态肉毒素, 额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (เกาหลี Innotox)","d":"โบท็อกซ์ชนิดน้ำหนึ่งเดียวในโลก, เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Innotox nội địa)","d":"Botox dạng lỏng duy nhất trên thế giới , Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (수입 제오민)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Xeomin (Imported)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"ja":{"e":"表情じわボトックスの1部位（海外製のゼオミン）","d":"おでこ・眉間・目元・鼻柱・鼻背・顎の梅干しじわから1部位選択"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（进口 Xeomin）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย 1 บริเวณ (นำเข้า Xeomin)","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / สันจมูก / ดั้งจมูก / คาง"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Xeomin nhập khẩu)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"사각턱보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox – Hutox (Korean)"},"ja":{"e":"エラボトックス（韓国製のハイトックス）"},"zh":{"e":"方下巴肉毒素（韩国产 HiTox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Hitox)"},"vi":{"e":"Botox hàm vuông (Haitox nội địa)"}}},{"k":"사각턱보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox Coretox (Korean)"},"ja":{"e":"エラボトックス（韓国製のコアトックス）"},"zh":{"e":"方下巴肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Coretox)"},"vi":{"e":"Botox hàm vuông (Coretox nội địa)"}}},{"k":"사각턱보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox Innotox (Korean)"},"ja":{"e":"エラボトックス（韓国製のイノトックス）"},"zh":{"e":"方下巴肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์กราม (เกาหลี Innotox)"},"vi":{"e":"Botox hàm vuông (Innotox nội địa)"}}},{"k":"사각턱보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox Xeomin (Imported)"},"ja":{"e":"エラボトックス（海外製のゼオミン）"},"zh":{"e":"方下巴肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์กราม (นำเข้า Xeomin)"},"vi":{"e":"Botox hàm vuông (Xeomin nhập khẩu)"}}},{"k":"풀페이스 스킨보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox – Hutox (Korean)"},"ja":{"e":"フルフェイススキンボトックス（韓国製のハイトックス）"},"zh":{"e":"全脸皮肤肉毒素（韩国产 HiTox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Hitox)"},"vi":{"e":"Skin Botox toàn mặt (Haitox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Coretox (Korean)"},"ja":{"e":"フルフェイススキンボトックス（韓国製のコアトックス）"},"zh":{"e":"全脸皮肤肉毒素（韩国产 Coretox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Coretox)"},"vi":{"e":"Skin Botox toàn mặt (Coretox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Innotox (Korean)"},"ja":{"e":"フルフェイススキンボトックス（韓国製のイノトックス）"},"zh":{"e":"全脸皮肤肉毒素（韩国产 Innotox）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (เกาหลี Innotox)"},"vi":{"e":"Skin Botox toàn mặt (Innotox nội địa)"}}},{"k":"풀페이스 스킨보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Xeomin (Imported)"},"ja":{"e":"フルフェイススキンボトックス（海外製のゼオミン）"},"zh":{"e":"全脸皮肤肉毒素（进口 Xeomin）"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า (นำเข้า Xeomin)"},"vi":{"e":"Skin Botox toàn mặt (Xeomin nhập khẩu)"}}},{"k":"승모근/종아리보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox – Hutox (Korean)"},"ja":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のハイトックス）"},"zh":{"e":"斜方肌/小腿肉毒素（韩国产 HiTox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Hitox)"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Haitox nội địa)"}}},{"k":"승모근/종아리보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Coretox (Korean)"},"ja":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のコアトックス）"},"zh":{"e":"斜方肌/小腿肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Coretox)"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Coretox nội địa)"}}},{"k":"승모근/종아리보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Innotox (Korean)"},"ja":{"e":"僧帽筋・ふくらはぎボトックス（韓国製のイノトックス）"},"zh":{"e":"斜方肌/小腿肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (เกาหลี Innotox)"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Innotox nội địa)"}}},{"k":"승모근/종아리보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Xeomin (Imported)"},"ja":{"e":"（海外製のゼオミン）"},"zh":{"e":"斜方肌/小腿肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์บ่า / น่อง (นำเข้า Xeomin)"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Xeomin nhập khẩu)"}}},{"k":"다한증보톡스 (국산 하이톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox – Hutox (Korean)"},"ja":{"e":"多汗症ボトックス（韓国製のハイトックス）"},"zh":{"e":"多汗症肉毒素（韩国产 HiTox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Hitox)"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Haitox nội địa)"}}},{"k":"다한증보톡스 (국산 코어톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Coretox (Korean)"},"ja":{"e":"多汗症ボトックス（韓国製のコアトックス）"},"zh":{"e":"多汗症肉毒素（韩国产 Coretox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Coretox)"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Coretox nội địa)"}}},{"k":"다한증보톡스 (국산 이노톡스)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Innotox (Korean)"},"ja":{"e":"多汗症ボトックス（韓国製のイノトックス）"},"zh":{"e":"多汗症肉毒素（韩国产 Innotox）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (เกาหลี Innotox)"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Innotox nội địa)"}}},{"k":"다한증보톡스 (수입 제오민)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Xeomin (Imported)"},"ja":{"e":"多汗症ボトックス（海外製のゼオミン）"},"zh":{"e":"多汗症肉毒素（进口 Xeomin）"},"th":{"e":"โบท็อกซ์ลดเหงื่อ (นำเข้า Xeomin)"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Xeomin nhập khẩu)"}}},{"k":"하관 대용량 양악주사","c":"보톡스/윤곽주사","t":{"en":{"e":"Lower Face Contouring Injection"},"ja":{"e":"フェイスラインの大容量輪郭注射"},"zh":{"e":"下脸部大容量双颌针"},"th":{"e":"ฉีดหน้าเรียว แบบปริมาณมาก"},"vi":{"e":"Tiêm hàm dưới liều dung lượng cao"}}},{"k":"코조각주사","kd":"-","c":"보톡스/윤곽주사","t":{"en":{"e":"Nose Contouring Injection","d":"-"},"ja":{"e":"鼻彫刻注射","d":"-"},"zh":{"e":"鼻部雕刻针","d":"-"},"th":{"e":"ฉีดลดปีกจมูก","d":"-"},"vi":{"e":"Tiêm tạo dáng mũi"}}},{"k":"볼륨필러 (국산 아띠에르) 1cc","c":"필러","t":{"en":{"e":"Volume Filler 1cc – Attiére (Korean)"},"ja":{"e":"ボリュームアップフィラー（韓国製のアティエル）1㏄"},"zh":{"e":"丰盈填充（韩国产 Atiere）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (เกาหลี Attiér) 1cc"},"vi":{"e":"Filler tạo phồng (Atiere nội địa) 1cc"}}},{"k":"볼륨필러 (국산 순수필) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Soonsu Fill 1cc (Korean)"},"ja":{"e":"ボリュームアップフィラー（韓国製のスンスフィル）1㏄"},"zh":{"e":"丰盈填充（韩国产 Sunsu Fill）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (เกาหลี Soonsu Fill) 1cc"},"vi":{"e":"Filler tạo phồng (Soonsoo Fill nội địa) 1cc"}}},{"k":"볼륨필러 (수입 아말리안) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Amalian 1cc (Imported)"},"ja":{"e":"ボリュームアップフィラー（海外製のアマリアン）1㏄"},"zh":{"e":"丰盈填充（进口 Amalian）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (นำเข้า Amalian) 1cc"},"vi":{"e":"Filler tạo phồng (Amalian nhập khẩu) 1cc"}}},{"k":"볼륨필러 (수입 쥬비덤) 1cc","c":"필러","t":{"en":{"e":"Volume Filler Juvederm 1cc (Imported)"},"ja":{"e":"ボリュームアップフィラー（海外製のジュビダーム）1㏄"},"zh":{"e":"丰盈填充（进口 Juvederm）1cc"},"th":{"e":"ฟิลเลอร์เพิ่มวอลุ่ม (นำเข้า Juvederm) 1cc"},"vi":{"e":"Filler tạo phồng (Juvederm nhập khẩu) 1cc"}}},{"k":"입술필러 무제한 (국산 아띠에르)","kd":"입꼬리 포함","c":"필러","t":{"en":{"e":"Unlimited Lip Filler – Attiére (Korean)","d":"Includes Mouth Corner Lift"},"ja":{"e":"唇フィラー無制限（韓国製のアティエル）","d":"口角を含む"},"zh":{"e":"嘴唇填充 不限容量（韩国产 Atiere）","d":"包含嘴角"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc (เกาหลี Attiér)","d":"(รวมยกมุมปาก)"},"vi":{"e":"Filler môi không giới hạn (Atiere nội địa)","d":"Bao gồm cả khóe miệng"}}},{"k":"입술필러 무제한 (국산 순수필)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Soonsu Fill (Korean)"},"ja":{"e":"唇フィラー無制限（韓国製のスンスフィル）1㏄"},"zh":{"e":"嘴唇填充 不限容量（韩国产 Sunsu Fill）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc(เกาหลี Soonsu Fill)"},"vi":{"e":"Filler môi không giới hạn (Soonsoo Fill nội địa)"}}},{"k":"입술필러 무제한 (수입 아말리안)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Amalian (Imported)"},"ja":{"e":"唇フィラー無制限（海外製のアマリアン）1㏄"},"zh":{"e":"嘴唇填充 不限容量（进口 Amalian）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc(นำเข้า Amalian)"},"vi":{"e":"Filler môi không giới hạn (Amalian nhập khẩu)"}}},{"k":"입술필러 무제한 (수입 레스틸렌키스)","c":"필러","t":{"en":{"e":"Unlimited Lip Filler Restylane Kysse (Imported)"},"ja":{"e":"唇フィラー無制限（海外製のレスチレンキス）"},"zh":{"e":"嘴唇填充 不限容量（进口 Restylane Kysse）"},"th":{"e":"ฟิลเลอร์ปากไม่จำกัด cc (นำเข้า Restylane Kysse)"},"vi":{"e":"Filler môi không giới hạn (Restylane Kysse nhập khẩu)"}}},{"k":"귀필러 (국산 아띠에르) 1cc","c":"필러","t":{"en":{"e":"Ear Filler 1cc – Attiére (Korean)"},"ja":{"e":"立ち耳フィラー（韓国製のアティエル）1㏄"},"zh":{"e":"耳部填充（韩国产 Atiere）1cc"},"th":{"e":"ฟิลเลอร์หู (เกาหลี Attiér) 1cc"},"vi":{"e":"Filler tai (Atiere nội địa) 1cc"}}},{"k":"귀필러 (수입 쥬비덤) 1cc","c":"필러","t":{"en":{"e":"Ear Filler Juvederm 1cc (Imported)"},"ja":{"e":"立ち耳フィラー（海外製のジュビダーム）1㏄"},"zh":{"e":"耳部填充（进口 Juvederm）1cc"},"th":{"e":"ฟิลเลอร์หู (นำเข้า Juvederm) 1cc"},"vi":{"e":"Filler tai (Juvederm nhập khẩu) 1cc"}}},{"k":"눈밑 전용필러 (수입 레스틸렌 아이라이트)","c":"필러","t":{"en":{"e":"Under-Eye Filler – Restylane Eyelight (Imported)"},"ja":{"e":"目の下の専用フィラー（海外製のレスチレンアイライト）"},"zh":{"e":"眼下专用填充（进口 Restylane Eyelight）"},"th":{"e":"ฟิลเลอร์สำหรับใต้ตาใต้ตา (นำเข้า Restylane Eyelight)"},"vi":{"e":"Filler chuyên dụng vùng dưới mắt (nhập khẩu Restylane Eye Light)"}}},{"k":"바디필러 1cc (국산)","c":"필러","t":{"en":{"e":"Body Filler 1cc (Korean)"},"ja":{"e":"ボディフィラー1㏄（韓国製）"},"zh":{"e":"身体填充 1cc（韩国产）"},"th":{"e":"บอดี้ฟิลเลอร์ 1cc (เกาหลี)"},"vi":{"e":"Filler body 1cc (nội địa)"}}},{"k":"리니어지 100샷","c":"리프팅","t":{"en":{"e":"LinearZ 100 Shots"},"ja":{"e":"リニアージ100ショット"},"zh":{"e":"LinearZ 100发"},"th":{"e":"ลิเนียร์จี 100 ช็อต"},"vi":{"e":"LinearZ 100 shot"}}},{"k":"인모드fx 풀페이스","c":"리프팅","t":{"en":{"e":"InMode FX Full Face"},"ja":{"e":"インモードfxフルフェイス"},"zh":{"e":"InMode FX 全脸"},"th":{"e":"อินโหมด FX ทั่วหน้า"},"vi":{"e":"InMode FX full face"}}},{"k":"인모드forma 풀페이스","c":"리프팅","t":{"en":{"e":"InMode FORMA Full Face"},"ja":{"e":"インモードformaフルフェイス"},"zh":{"e":"InMode Forma 全脸"},"th":{"e":"อินโหมด Forma ทั่วหน้า"},"vi":{"e":"InMode Forma full face"}}},{"k":"덴서티 100샷","c":"리프팅","t":{"en":{"e":"DENSITY 100 Shots"},"ja":{"e":"デンシティ100ショット"},"zh":{"e":"Density 100发"},"th":{"e":"เดนซิตี้ 100 ช็อต"},"vi":{"e":"Density 100 shot"}}},{"k":"포텐자 DDR 500샷","c":"리프팅","t":{"en":{"e":"Potenza DDR 500 Shots"},"ja":{"e":"ポテンツァDDR500ショット"},"zh":{"e":"Potenza DDR 500发"},"th":{"e":"โพเทนซ่า DDR 500 ช็อต"},"vi":{"e":"Potenza DDR 500 shot"}}},{"k":"포텐자 DIA 500샷","c":"리프팅","t":{"en":{"e":"Potenza DIA 500 Shots"},"ja":{"e":"ポテンツァDIA500ショット"},"zh":{"e":"Potenza DIA 500发"},"th":{"e":"โพเทนซ่า DIA 500 ช็อต"},"vi":{"e":"Potenza DIA 500 shot"}}},{"k":"울쎄라피프라임 100샷","c":"리프팅","t":{"en":{"e":"Ultherapy Prime 100 Shots"},"ja":{"e":"ウルセラプライム100ショット"},"zh":{"e":"Ultherapy Prime 100发"},"th":{"e":"อัลเทอราไพรม์ 100 ช็อต"},"vi":{"e":"Ultherapy Prime 100 shot"}}},{"k":"써마지FLX 300샷","c":"리프팅","t":{"en":{"e":"Thermage FLX 300 Shots"},"ja":{"e":"サーマクールFLX 300ショット"},"zh":{"e":"Thermage FLX 300发"},"th":{"e":"เทอร์มาจ FLX 300 ช็อต"},"vi":{"e":"Thermage FLX 300 shot"},"zh-TW":{"e":"第四代鳳凰電波 300發"}}},{"k":"물광주사 하이주 2cc","c":"스킨부스터","t":{"en":{"e":"Hycoox Skin Booster Injection 2cc"},"ja":{"e":"水光注射ハイジュ2㏄"},"zh":{"e":"水光针 Hyju 2cc"},"th":{"e":"ไฮจู 2cc"},"vi":{"e":"Tiêm căng bóng da High Inj 2cc"}}},{"k":"릴리이드M 5cc","c":"스킨부스터","t":{"en":{"e":"Lilied M 5cc"},"ja":{"e":"リリイドM5㏄"},"zh":{"e":"Lilied M 5cc"},"th":{"e":"Liliid M 5cc"},"vi":{"e":"Lilied M 5cc"}}},{"k":"리즈네 2cc","c":"스킨부스터","t":{"en":{"e":"Lizne 2cc"},"ja":{"e":"リズネ2㏄"},"zh":{"e":"Lizne 2cc"},"th":{"e":"Lizne 2cc"},"vi":{"e":"Lizne 2cc"}}},{"k":"아이 리쥬란 1cc","kd":"눈밑 전용","c":"스킨부스터","t":{"en":{"e":"Eye Rejuran 1cc","d":"For Under-Eye Area Only"},"ja":{"e":"リジュランアイ1㏄","d":"目の下専用"},"zh":{"e":"眼部 Rejuran 1cc","d":"眼底专用"},"th":{"e":"Eye Rejuran 1cc","d":"(สำหรับใต้ตาโดยเฉพาะ)"},"vi":{"e":"Rejuran Eye 1cc","d":"Chuyên dùng cho vùng mắt"}}},{"k":"리쥬란HB 1cc","c":"스킨부스터","t":{"en":{"e":"Rejuran HB 1cc"},"ja":{"e":"リジュランHB 1㏄"},"zh":{"e":"Rejuran HB 1cc"},"th":{"e":"Rejuran HB 1cc"},"vi":{"e":"Rejuran HB 1cc"}}},{"k":"리쥬란S 1cc","c":"스킨부스터","t":{"en":{"e":"Rejuran S 1cc"},"ja":{"e":"リジュランS1㏄"},"zh":{"e":"Rejuran S 1cc"},"th":{"e":"Rejuran S 1cc"},"vi":{"e":"Rejuran S 1cc"}}},{"k":"엑소좀 3D 네오스템 1병","kd":"항염효과 인증","c":"스킨부스터","t":{"en":{"e":"Exosome 3D NeoStem 1 Vial","d":"Certified Anti-Inflammatory Effect"},"ja":{"e":"ネオステム３Dエクソソーム1バイアル","d":"抗炎症作用認証"},"zh":{"e":"外泌体 3D Neostem 1瓶","d":"抗炎效果认证"},"th":{"e":"Exosome 3D Neostem 1 ขวด","d":"(ผ่านการรับรองเรื่องลดการอักเสบ)"},"vi":{"e":"Exosome 3D Neostem 1 lọ","d":"Chứng nhận hiệu quả kháng viêm"}}},{"k":"쥬베룩 스킨부스터 4cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Skin Booster 4cc"},"ja":{"e":"ジュベルックスキンブースター4㏄"},"zh":{"e":"Juvelook 水光针 4cc"},"th":{"e":"Juvelook Skin Booster 4cc"},"vi":{"e":"Juvelook Skin Booster 4cc"},"zh-TW":{"e":"Juvelook 水光針 4cc"}}},{"k":"쥬베룩 아이 1병","kd":"눈밑 전용","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial","d":"For Under-Eye Area Only"},"ja":{"e":"ジュベルックアイ1バイアル","d":"目の下専用"},"zh":{"e":"Juvelook 眼部 1瓶","d":"眼底专用"},"th":{"e":"Juvelook Eye 1 ขวด","d":"(สำหรับใต้ตาโดยเฉพาะ)"},"vi":{"e":"Juvelook Eye 1 lọ","d":"Chuyên dùng cho vùng mắt"},"zh-TW":{"e":"Juvelook 眼部 1瓶"}}},{"k":"쥬베룩 볼륨 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Volume 1 Vial"},"ja":{"e":"ジュベルックボリューム1バイアル"},"zh":{"e":"Juvelook 丰盈 1瓶"},"th":{"e":"Juvelook Volume 1 ขวด"},"vi":{"e":"Juvelook Volume 1 lọ"},"zh-TW":{"e":"Juvelook 豐盈 1瓶"}}},{"k":"쥬비덤 스킨바이브 1cc","c":"스킨부스터","t":{"en":{"e":"Juvederm Skinvive 1cc"},"ja":{"e":"スキンバイブbyジュビダーム"},"zh":{"e":"Juvederm Skinvive 1cc"},"th":{"e":"Juvederm Skinvive 1cc"},"vi":{"e":"Juvederm Skinvive 1cc"}}},{"k":"벨로테로 리바이브 1cc","c":"스킨부스터","t":{"en":{"e":"Belotero Revive 1cc"},"ja":{"e":"ベロテロリバイブ1㏄"},"zh":{"e":"Belotero Revive 1cc"},"th":{"e":"Belotero Revive 1cc"},"vi":{"e":"Belotero Revive 1cc"}}},{"k":"리투오 5cc","c":"스킨부스터","t":{"en":{"e":"LITUO 5cc"},"ja":{"e":"リトゥオ5㏄"},"zh":{"e":"Lituo 5cc"},"th":{"e":"Lituo 5cc"},"vi":{"e":"Re2O 5cc"}}},{"k":"셀르디엠 1시린지","c":"스킨부스터","t":{"en":{"e":"CellreDM 1 Syringe"},"ja":{"e":"セルディエム1シリンジ"},"zh":{"e":"Cell D.M 1支针剂"},"th":{"e":"Cell DM 1 Syringe"},"vi":{"e":"Cellredm 1 ống tiêm"}}},{"k":"스컬트라 1병","c":"스킨부스터","t":{"en":{"e":"Sculptra 1 Vial"},"ja":{"e":"スカルトラ1バイアル"},"zh":{"e":"Sculptra 1瓶"},"th":{"e":"Sculptra 1 ขวด"},"vi":{"e":"Sculptra 1 lọ"}}},{"k":"레디어스 1시린지","c":"스킨부스터","t":{"en":{"e":"Radiesse 1 Syringe"},"ja":{"e":"レディエッセ1シリンジ"},"zh":{"e":"Radiesse 1支针剂"},"th":{"e":"Radiesse 1 Syringe"},"vi":{"e":"Radiesse 1 ống tiêm"}}},{"k":"미라셀 줄기세포","c":"줄기세포","t":{"en":{"e":"Miracell Stem Cell Therapy"},"ja":{"e":"ミラセル幹細胞"},"zh":{"e":"Miracell 干细胞"},"th":{"e":"Miracell Stem Cell"},"vi":{"e":"Miracell tế bào gốc"}}},{"k":"울트라브이 콜라겐실 1줄","c":"실리프팅","t":{"en":{"e":"Ultra V Collagen Thread – 1 Thread"},"ja":{"e":"ウルトラブイのコラーゲン糸1本"},"zh":{"e":"Ultra V 胶原蛋白线 1根"},"th":{"e":"ไหมคอลลาเจน Ultra V 1 เส้น"},"vi":{"e":"Sợi chỉ collagen Ultra-V 1 sợi"}}},{"k":"실리프팅 1줄","kd":"심부볼,팔자,턱선 리프팅","c":"실리프팅","t":{"en":{"e":"Thread Lifting – 1 Thread","d":"Deep Cheek, Nasolabial Fold & Jawline Lifting"},"ja":{"e":"糸リフト1本","d":"バッカルファット、ほうれい線、 顎のラインリフトアップ"},"zh":{"e":"线雕 1根","d":"深层脸颊脂肪、法令纹、下颌线提升"},"th":{"e":"ร้อยไหม 1 เส้น","d":"(ยกกระพุ้งแก้ม ร่องแก้ม กรอบหน้า)"},"vi":{"e":"Căng chỉ 1 sợi","d":"Nâng cơ má hóp, rãnh cười, đường viền hàm"}}},{"k":"잼버실 2줄","kd":"팔자,볼,인디언주름 볼륨","c":"실리프팅","t":{"en":{"e":"Jamber Thread Lift – 2 Threads","d":"Nasolabial Fold, Cheek & Indian Fold Volume Enhancement"},"ja":{"e":"ジャンバー糸2本","d":"ほうれい線、頬、インディアンラインボリューム"},"zh":{"e":"Jambersil 2根","d":"法令纹、脸颊、印第安纹容量填充"},"th":{"e":"ไหม Jamber 2 เส้น","d":"(เพิ่มวอลุ่มร่องแก้ม กระพุ้งแก้ม ร่องอินเดียน)"},"vi":{"e":"Chỉ Jamber 2 sợi","d":"Tăng thể tích rãnh cười, má và rãnh lệ có nếp nhăn"}}},{"k":"이마거상,폭스아이 실리프팅 1줄","c":"실리프팅","t":{"en":{"e":"Brow Lift / Fox Eye Thread Lift – 1 Thread"},"ja":{"e":"おでこリフト、フォックスアイ糸リフト1本"},"zh":{"e":"额头提升、Fox Eye线雕 1根"},"th":{"e":"ร้อยไหมยกหน้าผาก / Foxy Eyes 1 เส้น"},"vi":{"e":"Căng chỉ trán, mắt cáo 1 sợi"}}},{"k":"콧대주름 실리프팅 1줄","c":"실리프팅","t":{"en":{"e":"Nose Bridge Wrinkle Thread Lift – 1 Thread"},"ja":{"e":"バニーラインの糸リフト1本"},"zh":{"e":"鼻梁皱纹线雕 1根"},"th":{"e":"ร้อยไหมสันจมูก 1 เส้น"},"vi":{"e":"Nâng chỉ vùng nếp nhăn sống mũi 1 sợi"}}},{"k":"이중턱 근육묶기 실리프팅 2줄","c":"실리프팅","t":{"en":{"e":"Double Chin Muscle Fixing Thread Lift – 2 Threads"},"ja":{"e":"二十顎の筋肉縛りの糸リフト2本"},"zh":{"e":"双下巴肌肉固定线雕 2根"},"th":{"e":"ร้อยไหมเก็บเหนียง 2 เส้น"},"vi":{"e":"Căng chỉ nâng cơ thu gọn cằm hai ngấn 2 sợi"}}},{"k":"점,비립종,사마귀제거 1개","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Mole / Milia / Wart Removal – 1 Lesion"},"ja":{"e":"ほくろ、稗粒腫、いぼの除去1個"},"zh":{"e":"痣、脂肪粒、去跖疣 1个"},"th":{"e":"กำจัดไฝ / สิวหิน / หูด 1 จุด"},"vi":{"e":"Loại bỏ nốt ruồi, mụn kê, hoặc mụn cóc 1 nốt"}}},{"k":"듀얼토닝(브라이톤+제네시스) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Dual Toning (Brightone + Genesis) – 1 Session"},"ja":{"e":"デュアルトーニング（ブライトン＋ジェネシス）1回"},"zh":{"e":"双重嫩肤（Brightone + Genesis）1次"},"th":{"e":"ดูอัลโทนนิ่ง (Brightone + Genesis) 1 ครั้ง"},"vi":{"e":"Dual Toning (Brighton + Noble) 1 lần"}}},{"k":"피코토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Pico Toning – 1 Session"},"ja":{"e":"ピコトーニング1回"},"zh":{"e":"Pico嫩肤 1次"},"th":{"e":"พิโคโทนนิ่ง 1 ครั้ง"},"vi":{"e":"Pico Toning 1 lần"}}},{"k":"헐리우드 스펙트라토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Toning – 1 Session"},"ja":{"e":"ハリウッドスペクトラ1回"},"zh":{"e":"Hollywood Spectra嫩肤 1次"},"th":{"e":"Hollywood Spectra Toning 1 ครั้ง"},"vi":{"e":"Hollywood Spectra Toning 1 lần"}}},{"k":"엑셀브이 플러스(홍조/색소) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Excel V Plus (Redness/Pigmentation) – 1 Session"},"ja":{"e":"エクセルVプラス（紅潮・色素）１回"},"zh":{"e":"Excel V Plus（红血丝/色素）1次"},"th":{"e":"Excel V พลัส (รอยแดง / เม็ดสี) 1 ครั้ง"},"vi":{"e":"Excel V+ (điều trị mẩn đỏ/ sắc tố) 1 lần"}}},{"k":"풀페이스 혈관레이저 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Full Face Vascular Laser – 1 Session"},"ja":{"e":"フルフェイス血管レーザー１回"},"zh":{"e":"全脸血管激光 1次"},"th":{"e":"เลเซอร์เส้นเลือดทั่วหน้า 1 ครั้ง"},"vi":{"e":"Laser mạch máu toàn mặt 1 lần"}}},{"k":"포텐자(기미) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Potenza (Melasma) – 1 Session"},"ja":{"e":"ポテンツァ（シミ）1回"},"zh":{"e":"Potenza（黄褐斑）1次"},"th":{"e":"โพเทนซ่า ฝ้า 1 ครั้ง"},"vi":{"e":"Potenza (điều trị nám) 1 lần"}}},{"k":"얼굴전체 잡티레이저(클라리티) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Full Face Pigmentation Laser (Clarity) – 1 Session"},"ja":{"e":"顔全体のシミ・くすみレーザー（クラリティ）1回"},"zh":{"e":"全脸色斑激光（Clarity）1次"},"th":{"e":"เลเซอร์กำจัดจุดด่างดำทั่วหน้า (Clarity) 1 ครั้ง"},"vi":{"e":"Laser đốm sắc tố toàn mặt (Clarity) – 1 lần"}}},{"k":"리팟 흑자제거 1부위","kd":"1cm 이내","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Age Spot Removal – 1 Area","d":"1cm 이내 Within 1cm"},"ja":{"e":"リポットレーザー黒子除去1部位","d":"1cm 이내　1㎝以内"},"zh":{"e":"Reepot 黑斑去除 1部位","d":"1cm以内"},"th":{"e":"Lipot กำจัดกระแดด 1 บริเวณ","d":"(ไม่เกิน 1 ซม.)"},"vi":{"e":"Reepot loại bỏ đốm nâu 1 vùng","d":"Trong phạm vi 1cm"}}},{"k":"눈썹 문신제거 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Eyebrow Tattoo Removal – 1 Session"},"ja":{"e":"眉毛のアートメイク除去1回"},"zh":{"e":"眉毛纹身去除 1次"},"th":{"e":"ลบรอยสักคิ้ว 1 ครั้ง"},"vi":{"e":"Xóa xăm chân mày 1 lần"}}},{"k":"아이라인 문신제거 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Eyeline Tattoo Removal – 1 Session"},"ja":{"e":"アイラインのアートメイク除去1回"},"zh":{"e":"眼线纹身去除 1次"},"th":{"e":"ลบรอยสักอายไลเนอร์ 1 ครั้ง"},"vi":{"e":"Xóa xăm mí mắt 1 lần"}}},{"k":"흑백,컬러,레터링 문신제거","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Black & Gray / Color / Lettering Tattoo Removal"},"ja":{"e":"黒と白、カラー、レタリングタトゥー除去"},"zh":{"e":"黑白、彩色、文字纹身去除"},"th":{"e":"ลบรอยสักขาวดำ / สี / ตัวอักษร"},"vi":{"e":"Xóa xăm đơn sắc, màu hoặc lettering"}}},{"k":"애플토닝 1회","kd":"피지선 억제","c":"여드름/모공/흉터","t":{"en":{"e":"Apple Toning – 1 Session","d":"Sebaceous Gland Suppression"},"ja":{"e":"アップルトーニング","d":"皮脂腺の抑制"},"zh":{"e":"Apple嫩肤 1次","d":"抑制皮脂腺"},"th":{"e":"แอปเปิลโทนนิ่ง 1 ครั้ง","d":"(ช่วยลดการทำงานของต่อมไขมัน)"},"vi":{"e":"Apple Toning 1 lần","d":"Ức chế tuyến bã nhờn"}}},{"k":"포텐자(모공/여드름) 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza (Pores/Acne) – 1 Session"},"ja":{"e":"ポテンツァ（毛穴・ニキビ）1回"},"zh":{"e":"Potenza（毛孔/痘痘）1次"},"th":{"e":"โพเทนซ่า รูขุมขน / สิว 1 ครั้ง"},"vi":{"e":"Potenza (lỗ chân lông / mụn trứng cá) 1 lần"}}},{"k":"어븀프락셀 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Erbium Fraxel – 1 Session"},"ja":{"e":"エルビウムプラクショナルレーザー1回"},"zh":{"e":"Erbium Fraxel 1次"},"th":{"e":"เออร์เบียมแฟรกเซล 1 ครั้ง"},"vi":{"e":"Erbium Fraxel 1 lần"}}},{"k":"피코프락셀 MLA 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel MLA – 1 Session"},"ja":{"e":"ピコプラクショナルレーザー1回"},"zh":{"e":"Pico Fraxel MLA 1次"},"th":{"e":"พิโคแฟรกเซล MLA 1 ครั้ง"},"vi":{"e":"Pico Fraxel MLA 1 lần"}}},{"k":"남성 인중제모 1회","c":"제모","t":{"en":{"e":"Men’s Upper Lip Hair Removal – 1 Session"},"ja":{"e":"男性鼻下脱毛1回"},"zh":{"e":"男性 人中脱毛 1次"},"th":{"e":"กำจัดขนหนวด (ชาย) 1 ครั้ง"},"vi":{"e":"Triệt ria mép nam 1 lần"}}},{"k":"남성 인중,앞턱제모 1회","c":"제모","t":{"en":{"e":"Men’s Upper Lip & Chin Hair Removal – 1 Session"},"ja":{"e":"男性鼻下、顎脱毛1回"},"zh":{"e":"男性 人中、下巴前部脱毛 1次"},"th":{"e":"กำจัดขนหนวดและเครา (ชาย) 1 ครั้ง"},"vi":{"e":"Triệt ria mép và cằm trước nam 1 lần"}}},{"k":"남성 산타제모 1회","kd":"인중,턱수염,구렛나루,볼","c":"제모","t":{"en":{"e":"Men’s Santa Beard Hair Removal – 1 Session","d":"Upper Lip, Beard, Sideburns & Cheeks"},"ja":{"e":"男性サンタ脱毛1回","d":"鼻下、顎、もみあげ、頬"},"zh":{"e":"男性 Santa脱毛 1次","d":"人中、胡须、鬓角、脸颊"},"th":{"e":"กำจัดขนบริเวณหนวดและแก้ม 1 ครั้ง","d":"หนวด, เครา, จอน, แก้ม"},"vi":{"e":"Triệt râu quai nón nam 1 lần","d":"Ria mép, râu cằm, tóc mai, má"}}},{"k":"남성 수염전체제모 1회","kd":"인중,턱수염,구렛나루,볼,목라인","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal – 1 Session","d":"Upper Lip, Beard, Sideburns, Cheeks & Neck Line"},"ja":{"e":"男性ヒゲ全部位脱毛1回","d":"鼻下、顎、もみあげ、頬、首"},"zh":{"e":"男性 全胡须脱毛 1次","d":"人中、胡须、鬓角、脸颊、颈部线条"},"th":{"e":"กำจัดขนทั่วใบหน้าและลำคอ (ชาย) 1 ครั้ง","d":"หนวด, เครา, จอน, แก้ม, แนวลำคอ"},"vi":{"e":"Triệt toàn bộ râu nam 1 lần","d":"Ria mép, râu cằm, tóc mai, má, vùng cổ"}}},{"k":"여성 인중제모 1회","c":"제모","t":{"en":{"e":"Women’s Upper Lip Hair Removal – 1 Session"},"ja":{"e":"女性鼻下脱毛1回"},"zh":{"e":"女性 人中脱毛 1次"},"th":{"e":"กำจัดขนหนวด (หญิง) 1 ครั้ง"},"vi":{"e":"Triệt ria mép nữ 1 lần"}}},{"k":"여성 겨드랑이제모 1회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal – 1 Session"},"ja":{"e":"女性両ワキ脱毛1回"},"zh":{"e":"女性 腋下脱毛 1次"},"th":{"e":"กำจัดขนรักแร้ (หญิง) 1 ครั้ง"},"vi":{"e":"Triệt lông nách nữ 1 lần"}}},{"k":"여성 종아리제모 1회","kd":"무릎포함","c":"제모","t":{"en":{"e":"Women’s Lower Leg Hair Removal – 1 Session","d":"Includes Knees"},"ja":{"e":"女性ふくらはぎ脱毛1回","d":"膝を含む"},"zh":{"e":"女性 小腿脱毛 1次","d":"包含膝盖"},"th":{"e":"กำจัดขนหน้าแข้ง (หญิง) 1 ครั้ง","d":"รวมเข่า"},"vi":{"e":"Triệt lông bắp chân nữ 1 lần","d":"Bao gồm cả đầu gối"}}},{"k":"바디 리니어지 100샷","c":"비만/다이어트","t":{"en":{"e":"Body LinearZ 100 Shots"},"ja":{"e":"ボディリニアージ100ショット"},"zh":{"e":"Body LinearZ 100发"},"th":{"e":"Body LinearZ 100 ช็อต"},"vi":{"e":"Body Linearz 100 shot"}}},{"k":"지방분해주사 용량무제한 1회","c":"비만/다이어트","t":{"en":{"e":"Unlimited Fat-Dissolving Injection – 1 Session"},"ja":{"e":"脂肪溶解注射・容量無制限1回"},"zh":{"e":"溶脂针 不限容量 1次"},"th":{"e":"ฉีดสลายไขมัน (ไม่จำกัดปริมาณ) 1 ครั้ง"},"vi":{"e":"Tiêm tan mỡ, không giới hạn liều, 1 lần"}}},{"k":"지방추출주사 에스펫 1병","kd":"50cc","c":"비만/다이어트","t":{"en":{"e":"S-FAT Fat Extraction Injection – 1 Vial","d":"50cc"},"ja":{"e":"脂肪吸引注射エスファット1瓶","d":"50cc"},"zh":{"e":"抽脂针 S-Pet 1瓶","d":"50cc"},"th":{"e":"S-Pet ดูดไขมัน (1 ขวด)","d":"50cc"},"vi":{"e":"Tiêm chiết xuất mỡ S-pat 1 lọ","d":"50cc"}}},{"k":"위고비","c":"비만/다이어트","t":{"en":{"e":"Wegovy"},"ja":{"e":"ウゴービ"},"zh":{"e":"Wegovy"},"th":{"e":"Wegovy"},"vi":{"e":"Wegovy"}}},{"k":"보조개 한쪽","c":"성형수술","t":{"en":{"e":"Dimple Creation – One Side"},"ja":{"e":"えくぼ形成片側"},"zh":{"e":"酒窝（单侧）"},"th":{"e":"ทำลักยิ้ม (1 ข้าง)"},"vi":{"e":"Một bên lúm đồng tiền"}}},{"k":"풀페이스 지방이식","kd":"수면마취 가능","c":"성형수술","t":{"en":{"e":"Full Face Fat Grafting","d":"Sedation Anesthesia Available"},"ja":{"e":"フルフェイス脂肪注入","d":"睡眠麻酔可能"},"zh":{"e":"全脸脂肪移植","d":"可进行睡眠麻醉"},"th":{"e":"เติมไขมันทั่วใบหน้า (Fat Grafting)","d":"หลับด้วยยาสลบได้"},"vi":{"e":"Cấy mỡ toàn mặt","d":"Có thể gây mê ngủ"}}},{"k":"눈밑 지방제거+재배치+눈밑,앞볼 지방이식","kd":"수면마취 가능","c":"성형수술","t":{"en":{"e":"Under-Eye Fat Removal + Fat Repositioning + Under-Eye & Front Cheek Fat Grafting"},"ja":{"e":"目の下の脂肪取り＋再配置＋目の下、頬の脂肪注入"},"zh":{"e":"眼下脂肪去除+重新排列+眼下、前脸颊脂肪移植"},"th":{"e":"จัดเรียงไขมันใต้ตา + เติมไขมันใต้ตาและแก้ม"},"vi":{"e":"Lấy mỡ bọng mắt + tái phân bố + cấy mỡ vùng dưới mắt và má trước"}}},{"k":"미백/보습/진정/재생관리 1회","c":"성형수술","t":{"en":{"e":"Brightening / Hydration / Soothing / Regeneration Care – 1 Session"},"ja":{"e":"美白・保湿・鎮静・再生管理1回"},"zh":{"e":"美白/保湿/舒缓/再生护理 1次"},"th":{"e":"โปรแกรมกระจ่างใส/เติมความชุ่มชื้น/ปลอบประโลมผิว 1 ครั้ง"},"vi":{"e":"Chăm sóc làm trắng/giữ ẩm/làm dịu/tái tạo 1 lần"}}},{"k":"아쿠아필+진정관리 1회","c":"성형수술","t":{"en":{"e":"Aqua Peel + Soothing Care – 1 Session"},"ja":{"e":"アクアピール＋鎮静管理1回"},"zh":{"e":"Aqua Peel + 舒缓护理 1次"},"th":{"e":"อาควาพีล + ทรีทเม้นท์ปลอบประโลมผิว 1 ครั้ง"},"vi":{"e":"Aqua Peel + chăm sóc làm dịu 1 lần"}}},{"k":"LDM 물방울 초음파관리 1회","c":"성형수술","t":{"en":{"e":"LDM Water-Drop Ultrasound Treatment – 1 Session"},"ja":{"e":"LDM水玉超音波管理1回"},"zh":{"e":"LDM 水滴超声波护理 1次"},"th":{"e":"LDM Water Drop Lifting 1 ครั้ง"},"vi":{"e":"Chăm sóc da bằng sóng siêu âm LDM Water Drop 1 lần"}}},{"k":"오투덤 산소테라피 1회","c":"성형수술","t":{"en":{"e":"O2Derm Oxygen Therapy – 1 Session"},"ja":{"e":"オーツーダム酸素テラピー1回"},"zh":{"e":"O2Derm 氧气疗法 1次"},"th":{"e":"O2derm Oxygen Therapy 1 ครั้ง"},"vi":{"e":"O2 Derm oxygen therapy 1 lần"}}},{"k":"4세대 라라필 1회","c":"성형수술","t":{"en":{"e":"4th Generation LHALALA Peel – 1 Session"},"ja":{"e":"第4世代のララピール1回"},"zh":{"e":"第四代 Lhala Peel 1次"},"th":{"e":"LHALA Peel (Gen 4) 1 ครั้ง"},"vi":{"e":"Lhala Peel thế hệ 4 – 1 lần"}}},{"k":"백옥/신데렐라주사 1회","c":"성형수술","t":{"en":{"e":"Glutathione / Cinderella IV Drip – 1 Session"},"ja":{"e":"白玉・シンデレラ注射1回"},"zh":{"e":"白玉/灰姑娘针 1次"},"th":{"e":"ฉีดกลูต้า/ฉีดซินเดอเรลล่า 1 ครั้ง"},"vi":{"e":"Tiêm Bạch Ngọc hoặc Cinderella 1 lần"}}},{"k":"항노화주사 1회","kd":"피로 회복,항산화","c":"성형수술","t":{"en":{"e":"Anti-Aging IV Drip – 1 Session","d":"Fatigue Recovery & Antioxidant Effect"},"ja":{"e":"抗老化注射1回","d":"疲労回復、抗酸化"},"zh":{"e":"抗衰老针 1次","d":"疲劳恢复、抗氧化"},"th":{"e":"ฉีดชะลอวัย 1 ครั้ง","d":"ฟื้นฟูร่างกาย, ต้านอนุมูลอิสระ"},"vi":{"e":"Tiêm chống lão hóa 1 lần","d":"Đỡ mệt mỏi, chống oxy hóa"}}},{"k":"상쾌한주사 1회","kd":"숙취 해소,피로 회복","c":"성형수술","t":{"en":{"e":"Energy Recovery IV Drip – 1 Session","d":"Hangover Relief & Fatigue Recovery"},"ja":{"e":"爽快注射1回","d":"二日酔い解消、疲労回復"},"zh":{"e":"清爽针 1次","d":"解酒、疲劳恢复"},"th":{"e":"ฉีดแก้แฮงค์ 1 ครั้ง","d":"แก้แฮงค์, ฟื้นฟูร่างกาย"},"vi":{"e":"Tiêm phục hồi sảng khoái 1 lần","d":"Giải rượu, hồi phục mệt mỏi"}}},{"k":"신비주사(신데렐라+비타민) 1회","kd":"노화 예방,체지방 감소,활성산소 제거,면역강화","c":"성형수술","t":{"en":{"e":"Shinbi IV Drip (Cinderella + Vitamins) – 1 Session","d":"Anti-Aging, Body Fat Reduction, Free Radical Removal & Immune Support"},"ja":{"e":"シンビ注射（シンデレラ＋ビタミン）1回","d":"老化予防、体脂肪減少、活性酸素除去、免疫強化"},"zh":{"e":"神秘针（灰姑娘+维他命）1次","d":"抗衰老、减少体脂、清除活性氧、增强免疫力"},"th":{"e":"ฉีด Sinbi (ฉีดซินเดอเรลล่า+ฉีดวิตามิน) 1 ครั้ง","d":"คุมน้ำหนัก, เสริมภูมิคุ้มกัน"},"vi":{"e":"Tiêm Shinbi (Cinderella + vitamin) 1 lần","d":"Ngăn ngừa lão hóa, giảm mỡ cơ thể, loại bỏ oxy hoạt tính, tăng cường miễn dịch"}}},{"k":"오로라주사(백옥+비타민) 1회","kd":"노화 예방,피부톤 개선,간 해독,활성산소 제거","c":"성형수술","t":{"en":{"e":"Aurora IV Drip (Glutathione + Vitamins) – 1 Session","d":"Anti-Aging, Skin Tone Improvement, Liver Detox & Free Radical Removal"},"ja":{"e":"オーロラ注射（白玉＋ビタミン）1回","d":"老化予防、肌のトーンアップ、肝臓解毒、活性酸素除去"},"zh":{"e":"极光针（白玉+维他命）1次","d":"抗衰老、改善肤色、肝脏解毒、清除活性氧"},"th":{"e":"ฉีด Aurora(ฉีดกลูต้า+ฉีดวิตามิน) 1 ครั้ง","d":"ดีท็อกซ์ตับ, ปรับสีผิว"},"vi":{"e":"Tiêm Aurora (Bạch Ngọc + vitamin) 1 lần","d":"Ngăn ngừa lão hóa, cải thiện tone da, thải độc gan, loại bỏ oxy hoạt tính"}}},{"k":"소녀주사(태반+백옥) 1회","kd":"갱년기증상 완화,피부톤 개선,저속노화","c":"성형수술","t":{"en":{"e":"Girl IV Drip (Placenta + Glutathione) – 1 Session","d":"Menopause Symptom Relief, Skin Tone Improvement & Slow Aging Support"},"ja":{"e":"少女注射（プラセンタ＋白玉）1回","d":"更年期症状改善、肌のトーンアップ、スローエイジング"},"zh":{"e":"少女针（胎盘素+白玉）1次","d":"缓解更年期症状、改善肤色、延缓衰老"},"th":{"e":"ฉีด So-Nyeo(ฉีดรกเด็ก+ฉีดกลูต้า) 1 ครั้ง","d":"ปรับสมดุลวัยทอง, ชะลอวัย"},"vi":{"e":"Tiêm trẻ hóa (nhau thai + Bạch Ngọc) 1 lần","d":"Giảm triệu chứng mãn kinh, cải thiện tone da, làm chậm lão hóa"}}},{"k":"표정주름보톡스 1부위 (국산 이노톡스)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) Innotox (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Innotox）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Innotox nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (국산 제테마)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (1 Area) JETEMA (Korean)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（韩国产 Jetema）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Jetema nội địa)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"표정주름보톡스 1부위 (수입 디스포트)","kd":"이마/미간/눈가/콧대/콧등/자갈턱 중 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Dysport (Imported)","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Chin"},"zh":{"e":"表情肌肉皱纹肉毒素 1部位（进口 Dysport）","d":"额头/眉间/眼周/鼻梁/鼻背/酒窝下巴中任选1"},"vi":{"e":"Botox nếp nhăn vùng mặt – 1 vùng (Dysport nhập khẩu)","d":"Chọn 1 trong các vùng trán / giữa chân mày / đuôi mắt / sống mũi / lưng mũi / cằm đá cuội"}}},{"k":"사각턱보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox JETEMA (Korean)"},"zh":{"e":"方下巴肉毒素（韩国产 Jetema）"},"vi":{"e":"Botox hàm vuông(Jetema nội địa)"}}},{"k":"사각턱보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Jawline Botox / Masseter Botox Dysport (Imported)"},"zh":{"e":"方下巴肉毒素（进口 Dysport）"},"vi":{"e":"Botox hàm vuông(Dysport nhập khẩu)"}}},{"k":"풀페이스 스킨보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox JETEMA (Korean)"},"zh":{"e":"全脸皮肤肉毒素（韩国产 Jetema）"},"vi":{"e":"Skin Botox toàn mặt (Jetema nội địa)"}}},{"k":"풀페이스 스킨보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Full Face Skin Botox Dysport (Imported)"},"zh":{"e":"全脸皮肤肉毒素（进口 Dysport）"},"vi":{"e":"Skin Botox toàn mặt (Dysport nhập khẩu)"}}},{"k":"승모근/종아리보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox JETEMA (Korean)"},"zh":{"e":"斜方肌/小腿肉毒素（韩国产 Jetema）"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Jetema nội địa)"}}},{"k":"승모근/종아리보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Trapezius / Calf Botox Dysport (Imported)"},"zh":{"e":"斜方肌/小腿肉毒素（进口 Dysport）"},"vi":{"e":"Botox cơ thang/cơ bắp chân (Dysport nhập khẩu)"}}},{"k":"다한증보톡스 (국산 제테마)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox JETEMA (Korean)"},"zh":{"e":"多汗症肉毒素（韩国产 Jetema）"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Jetema nội địa)"}}},{"k":"다한증보톡스 (수입 디스포트)","c":"보톡스/윤곽주사","t":{"en":{"e":"Hyperhidrosis Botox Dysport (Imported)"},"zh":{"e":"多汗症肉毒素（进口 Dysport）"},"vi":{"e":"Botox trị chứng tiết nhiều mồ hôi (Dysport nhập khẩu)"}}},{"k":"볼륨필러 (국산 리쥬비엘) 1cc","c":"필러","t":{"en":{"e":"Volume Fille Rejuviel (Korean) 1cc"},"zh":{"e":"丰盈填充（韩国产 Rejuviel）1cc"},"vi":{"e":"Filler tạo phồng (Rejuviel nội địa) 1cc"}}},{"k":"입술필러 무제한 (국산 리쥬비엘)","c":"필러","t":{"en":{"e":"Unlimited Lip Fille Rejuviel (Korean)"},"zh":{"e":"嘴唇填充 不限容量（韩国产 Rejuviel）"},"vi":{"e":"Filler môi không giới hạn (Rejuviel nội địa) 1cc"}}},{"k":"눈밑 전용필러 (수입 벨로테로)","c":"필러","t":{"en":{"e":"Under-Eye Filler – Belotero (Imported)"},"zh":{"e":"眼下专用填充（进口 Belotero）"},"vi":{"e":"Filler chuyên dùng dưới mắt (Belotero nhập khẩu)"}}},{"k":"온다 1만줄","c":"리프팅","t":{"en":{"e":"ONDA 10,000 Shots"},"zh":{"e":"ONDA 1万发"},"vi":{"e":"Onda 10.000 line"}}},{"k":"슈링크 100샷","c":"리프팅","t":{"en":{"e":"Shurink 100 Shots"},"zh":{"e":"Shurink 100发"},"vi":{"e":"Shurink 100 shot"},"zh-TW":{"e":"Shurink 100發"}}},{"k":"볼뉴머 100샷","c":"리프팅","t":{"en":{"e":"Volnewmer 100 Shots"},"zh":{"e":"Volnewmer 100发"},"vi":{"e":"Volnewmer 100 shot"}}},{"k":"울쎄라피 프라임 100샷","c":"리프팅","t":{"en":{"e":"Ultherapy Prime 100 Shots"},"zh":{"e":"Ultherapy Prime 100发"},"vi":{"e":"Ultherapy Prime 100 shot"}}},{"k":"물광주사 2cc","c":"스킨부스터","t":{"en":{"e":"Skin Booster Injection 2cc"},"zh":{"e":"水光针 2cc"},"vi":{"e":"Tiêm căng bóng 2cc"}}},{"k":"아이 리쥬란 1cc","kd":"눈가 전용","c":"스킨부스터","t":{"en":{"e":"Eye Rejuran 1cc","d":"For Eye Area Only"},"zh":{"e":"眼部 Rejuran 1cc","d":"眼周专用"},"vi":{"e":"Rejuran Eye 1cc","d":"Chuyên dùng cho vùng mắt"}}},{"k":"리쥬란힐러2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran Healer 2cc"},"zh":{"e":"Rejuran Healer 2cc"},"vi":{"e":"Rejuran Healer 2cc"}}},{"k":"쥬베룩 아이 1병","kd":"눈가 전용","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial","d":"For Eye Area Only"},"zh":{"e":"Juvelook 眼部 1瓶","d":"眼周专用"},"vi":{"e":"Juvelook Eye 1 lọ","d":"Chuyên dùng cho vùng mắt"}}},{"k":"실리프팅 1줄","kd":"팔자,턱선,심부볼 리프팅","c":"실리프팅","t":{"en":{"e":"Thread Lifting – 1 Thread","d":"Nasolabial Fold / Jawline / Deep Cheek Lifting"},"zh":{"e":"线雕 1根","d":"法令纹、下颌线、深层脸颊脂肪提升"},"vi":{"e":"Căng chỉ 1 sợi","d":"Nâng cơ rãnh cười, đường viền hàm, má hóp"}}},{"k":"듀얼토닝(브라이톤+노블) 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Dual Toning (Brighton + Noble) – 1 Session"},"zh":{"e":"双重嫩肤（Brightone + Noble）1次"},"vi":{"e":"Dual Toning (Brighton + Noble) 1 lần"}}},{"k":"노블 잡티레이저 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Noble Pigmentation Laser – 1 Session"},"zh":{"e":"Noble 色斑激光 1次"},"vi":{"e":"Laser đốm sắc tố Noble 1 lần"}}},{"k":"프락셀 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Fraxel – 1 Session"},"zh":{"e":"Fraxel 1次"},"vi":{"e":"Fraxel 1 lần"}}},{"k":"바디 온다 1만줄","c":"비만/다이어트","t":{"en":{"e":"Body ONDA 10,000 Shots"},"zh":{"e":"Body ONDA 1万发"},"vi":{"e":"Onda cho body 10.000 shot"}}},{"k":"VPL 바디 지방분해주사 1회","kd":"500cc~1000c","c":"비만/다이어트","t":{"en":{"e":"VPL Body Fat-Dissolving Injection – 1 Session","d":"500cc~1000c"},"zh":{"e":"VPL 身体溶脂针 1次","d":"500cc~1000cc"},"vi":{"e":"Tiêm tan mỡ VPL body 1 lần","d":"500cc~1000c"}}},{"k":"백옥or 신데렐라주사 1회","c":"비만/다이어트","t":{"en":{"e":"Glutathione or Cinderella IV Drip – 1 Session"},"zh":{"e":"白玉or灰姑娘针 1次"},"vi":{"e":"Tiêm Bạch Ngọc hoặc Cinderella 1 lần"}}},{"k":"이마 보톡스 기본","kd":"이마 주름 개선 자연스러운 표정 유지","c":"보톡스/윤곽주사","t":{"en":{"e":"Basic Forehead Botox","d":"Forehead Wrinkle Improvement with Natural Facial Expressions"},"zh":{"e":"额头肉毒素基本款","d":"额头皱纹改善，自然表情维持"},"th":{"e":"โบท็อกซ์หน้าผาก","d":"ช่วยลดริ้วรอยหน้าผาก พร้อมคงความเป็นธรรมชาติของสีหน้า"},"vi":{"e":"Botox trán cơ bản","d":"Cải thiện nếp nhăn trán, giữ biểu cảm tự nhiên"}}},{"k":"하이톡스 표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Hutox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"zh":{"e":"HiTox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย Hitox","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vi":{"e":"Haitox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"하이톡스 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Hutox Jaw Botox / Masseter Botox"},"zh":{"e":"HiTox 方下巴"},"th":{"e":"โบท็อกซ์กราม Hitox"},"vi":{"e":"Haitox Cằm vuông"}}},{"k":"하이톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Hutox Skin Botox / Full Face"},"zh":{"e":"HiTox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า Hitox"},"vi":{"e":"Haitox Skin Botox/toàn mặt"}}},{"k":"하이톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Hutox Skin Botox / Half Face"},"zh":{"e":"HiTox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า Hitox"},"vi":{"e":"Haitox Skin Botox/nửa mặt"}}},{"k":"하이톡스 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Hutox Skin Botox / Partial Area"},"zh":{"e":"HiTox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด Hitox"},"vi":{"e":"Haitox Skin Botox/ một phần"}}},{"k":"코어톡스  표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"zh":{"e":"Coretox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย Coretox","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vi":{"e":"CoreTox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"코어톡스   사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Jaw Botox / Masseter Botox"},"zh":{"e":"Coretox 方下巴"},"th":{"e":"โบท็อกซ์กราม Coretox"},"vi":{"e":"CoreTox Cằm vuông"}}},{"k":"코어톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Full Face"},"zh":{"e":"Coretox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า Coretox"},"vi":{"e":"CoreTox Skin Botox/toàn mặt"}}},{"k":"코어톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Half Face"},"zh":{"e":"Coretox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า Coretox"},"vi":{"e":"CoreTox Skin Botox/nửa mặt"}}},{"k":"코어톡스 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Coretox Skin Botox / Partial Area"},"zh":{"e":"Coretox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด Coretox"},"vi":{"e":"CoreTox Skin Botox/ một phần"}}},{"k":"이노톡스 표정근육주름","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Dynamic Wrinkle Botox","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"zh":{"e":"Innotox 表情肌肉皱纹","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选11"},"th":{"e":"โบท็อกซ์ลดริ้วรอย Innotox","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vi":{"e":"Innotox cho nếp nhăn cơ biểu cảm","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"이노톡스 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Jaw Botox / Masseter Botox"},"zh":{"e":"Innotox 方下巴"},"th":{"e":"โบท็อกซ์กราม Innotox"},"vi":{"e":"Innotox Cằm vuông"}}},{"k":"이노톡스 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Full Face"},"zh":{"e":"Innotox 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า Innotox"},"vi":{"e":"Innotox Skin Botox/toàn mặt"}}},{"k":"이노톡스 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Half Face"},"zh":{"e":"Innotox 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า Innotox"},"vi":{"e":"Innotox Skin Botox/nửa mặt"}}},{"k":"이노톡스  스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Innotox Skin Botox / Partial Area"},"zh":{"e":"Innotox 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด Innotox"},"vi":{"e":"Innotox Skin Botox/ một phần"}}},{"k":"제오민 표정근육주름 1부위","kd":"이마 미간 눈가 눈밑 콧등 자갈턱 입꼬리 택1","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Dynamic Wrinkle Botox – 1 Area","d":"Choose 1: Forehead / Frown Lines / Crow’s Feet / Under-Eye Area / Bunny Lines / Pebble Chin / Mouth Corners"},"zh":{"e":"Xeomin 表情肌肉皱纹 1部位","d":"额头/眉间/眼周/眼下/鼻背/酒窝下巴/嘴角 任选1"},"th":{"e":"โบท็อกซ์ลดริ้วรอย Xeomin 1 จุด","d":"เลือก 1 บริเวณ: หน้าผาก / ระหว่างคิ้ว / หางตา / ใต้ตา / สันจมูก / คาง / ยกมุมปาก"},"vi":{"e":"Xeomin cho nếp nhăn cơ biểu cảm 1 vùng","d":"Chọn 1 vùng trán, giữa hai chân mày, khóe mắt, dưới mắt, sống mũi, cằm đá cuội, khóe miệng"}}},{"k":"제오민 사각턱","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Jaw Botox / Masseter Botox"},"zh":{"e":"Xeomin 方下巴"},"th":{"e":"โบท็อกซ์กราม Xeomin"},"vi":{"e":"Xeomin Cằm vuông"}}},{"k":"제오민 스킨보톡스/풀페이스","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Full Face"},"zh":{"e":"Xeomin 皮肤肉毒素/全脸"},"th":{"e":"สกินโบท็อกซ์ทั่วหน้า Xeomin"},"vi":{"e":"Xeomin Skin Botox/toàn mặt"}}},{"k":"제오민 스킨보톡스/하프","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Half Face"},"zh":{"e":"Xeomin 皮肤肉毒素/半脸"},"th":{"e":"สกินโบท็อกซ์ครึ่งหน้า Xeomin"},"vi":{"e":"Xeomin Skin Botox/nửa mặt"}}},{"k":"제오민 스킨보톡스/부분","c":"보톡스/윤곽주사","t":{"en":{"e":"Xeomin Skin Botox / Partial Area"},"zh":{"e":"Xeomin 皮肤肉毒素/局部"},"th":{"e":"สกินโบท็อกซ์เฉพาะจุด Xeomin"},"vi":{"e":"Xeomin Skin Botox/ một phần"}}},{"k":"V-FIT 주사 (윤곽주사) 1회","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection – 1 Session"},"zh":{"e":"V-FIT 注射（轮廓针）1次"},"th":{"e":"ฉีด V-FIT (ฉีดแฟต) 1 ครั้ง"},"vi":{"e":"Tiêm V-FIT (Tiêm thon gọn line khuôn mặt) 1 lần"}}},{"k":"V-FIT 주사 (윤곽주사) 3회","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection Contouring Injection – 3 Sessions"},"zh":{"e":"V-FIT 注射（轮廓针）3次"},"th":{"e":"ฉีด V-FIT (ฉีดแฟต) 3 ครั้ง"},"vi":{"e":"Tiêm V-FIT (Tiêm thon gọn line khuôn mặt) 3 lần"}}},{"k":"슈퍼 V-FIT 주사 (슈퍼윤곽주사)1회","c":"보톡스/윤곽주사","t":{"en":{"e":"Super V-FIT Contouring Injection – 1 Session"},"zh":{"e":"超级 V-FIT 注射（超级轮廓针）1次"},"th":{"e":"Super V-FIT ฉีด (ฉีดแฟตสูตร Super) 1 ครั้ง"},"vi":{"e":"Tiêm Super V-FIT (tiêm siêu thon gọn line khuôn mặt) 1 lần"}}},{"k":"슈퍼 V-FIT 주사 (슈퍼윤곽주사)3회","c":"보톡스/윤곽주사","t":{"en":{"e":"Super V-FIT Contouring Injection Super Contouring Injection – 3 Sessions"},"zh":{"e":"超级 V-FIT 注射（超级轮廓针）3次"},"th":{"e":"Super V-FIT ฉีด (ฉีดแฟตสูตร Super) 3 ครั้ง"},"vi":{"e":"Tiêm Super V-FIT (Tiêm siêu thon gọn line khuôn mặt) 3 lần"}}},{"k":"양악주사 1회","c":"보톡스/윤곽주사","t":{"en":{"e":"Double Jaw Contouring Injection – 1 Session"},"zh":{"e":"双颌针 1次"},"th":{"e":"ฉีดหน้าเรียว 1 ครั้ง"},"vi":{"e":"Tiêm cằm 1 lần"}}},{"k":"양악주사 3회","c":"보톡스/윤곽주사","t":{"en":{"e":"Double Jaw Contouring Injection - 3 Sessions"},"zh":{"e":"双颌针 3次"},"th":{"e":"ฉีดหน้าเรียว 3 ครั้ง"},"vi":{"e":"Tiêm cằm 3 lần"}}},{"k":"아띠에르(이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Attiére Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"zh":{"e":"Attire（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Attiér ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vi":{"e":"Atiere (trán, thái dương, má, rãnh cười, cằm lẹm) 1cc"}}},{"k":"아띠에르(코) 1CC","c":"필러","t":{"en":{"e":"Attiére Fille (Nose Filler) 1cc"},"zh":{"e":"Attire（鼻子）1CC"},"th":{"e":"Attiér (ฟิลเลอร์จมูก) 1cc"},"vi":{"e":"Atiere (Mũi ) 1CC"}}},{"k":"순수필 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Soonsu Fill (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"zh":{"e":"Pure Feel（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Soonsoo Fill ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vi":{"e":"Soonsoo Fill (trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"순수필 (코) 1CC","c":"필러","t":{"en":{"e":"Soonsu Fill( Nose Filler) 1cc"},"zh":{"e":"Pure Feel（鼻子）1CC"},"th":{"e":"Soonsoo Fill (ฟิลเลอร์จมูก) 1cc"},"vi":{"e":"Soonsoo Fill (Mũi ) 1CC"}}},{"k":"아말리안 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Amalian Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"zh":{"e":"Amalian（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Amalian ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vi":{"e":"Amalian (trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"아말리안  (코) 1CC","c":"필러","t":{"en":{"e":"Amalian Filler (Nose Filler) 1cc"},"zh":{"e":"Amalian（鼻子）1CC"},"th":{"e":"Amalian(ฟิลเลอร์จมูก) 1cc"},"vi":{"e":"Amalian  (Mũi ) 1CC"}}},{"k":"와이솔루션(이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Y Solution Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"zh":{"e":"Y Solution（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Y Solution ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vi":{"e":"Y Solution (trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"와이솔루션 (코) 1CC","c":"필러","t":{"en":{"e":"Y Solution Filler (Nose Filler) 1cc"},"zh":{"e":"Y Solution（鼻子）1CC"},"th":{"e":"Y Solution (ฟิลเลอร์จมูก) 1cc"},"vi":{"e":"Y Solution  (Mũi ) 1CC"}}},{"k":"쥬비덤 (이마,관자,볼,팔자,무턱) 1CC","c":"필러","t":{"en":{"e":"Juvederm Filler (Forehead/Temples/Cheeks/Nasolabial Folds/Chin) 1cc"},"zh":{"e":"Juvederm（额头、太阳穴、脸颊、法令纹、下巴后缩）1CC"},"th":{"e":"Juvederm ฟิลเลอร์ (หน้าผาก / ขมับ / แก้ม / ร่องแก้ม / คาง) 1cc"},"vi":{"e":"Juvederm (trán, thái dương, má, rãnh cười, cằm lẹm) 1CC"}}},{"k":"쥬비덤 (코) 1CC","c":"필러","t":{"en":{"e":"Juvederm Filler (Nose Filler) 1cc"},"zh":{"e":"Juvederm（鼻子）1CC"},"th":{"e":"Juvederm(ฟิลเลอร์จมูก) 1cc"},"vi":{"e":"Juvederm (Mũi ) 1CC"}}},{"k":"엘프요정 귀필러 1CC","c":"필러","t":{"en":{"e":"Elf Ear Filler 1cc"},"zh":{"e":"精灵耳填充 1CC"},"th":{"e":"ฟิลเลอร์หูเอลฟ์ 1cc"},"vi":{"e":"Filler tai Elf 1CC"}}},{"k":"엘프요정 귀필러 용량 무제한 귀필러(리터치포함)","c":"필러","t":{"en":{"e":"Elf Ear Fille Unlimited Ear Filler (Includes Touch-Up)"},"zh":{"e":"精灵耳填充 不限容量耳部填充（含补打）"},"th":{"e":"ฟิลเลอร์หูเอลฟ์ ฟิลเลอร์หูไม่จำกัดปริมาณ (รวมรีทัช)"},"vi":{"e":"Filler tai Elf Filler tai không giới hạn dung tích (bao gồm retouch)"}}},{"k":"셀르디엠 1병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 1 Vial"},"zh":{"e":"CellREDM 1瓶"},"th":{"e":"Cell DM 1 ขวด"},"vi":{"e":"CellREDM 1 lọ"}}},{"k":"셀르디엠 2병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 2 Vials"},"zh":{"e":"CellREDM 2瓶"},"th":{"e":"Cell DM 2 ขวด"},"vi":{"e":"CellREDM 2 lọ"}}},{"k":"셀르디엠 3병","c":"스킨부스터","t":{"en":{"e":"CellreDM – 3 Vials"},"zh":{"e":"CellREDM 3瓶"},"th":{"e":"Cell DM 3 ขวด"},"vi":{"e":"CellREDM 3 lọ"}}},{"k":"하이엔드 안티에이징 줄기세포 스킨주사 1회","c":"줄기세포","t":{"en":{"e":"High-End Anti-Aging Stem Cell Skin Injection – 1 Session"},"zh":{"e":"高端抗衰老干细胞水光针 1次"},"th":{"e":"ฉีด High-End Anti-Aging Stem Cell Skin 1 ครั้ง"},"vi":{"e":"Tiêm skin tế bào gốc chống lão hóa high-end 1 lần"}}},{"k":"하이앤드 안티에이징 줄기세포 스킨주사 1회+ 리투오 1병","c":"줄기세포","t":{"en":{"e":"High-End Anti-Aging Stem Cell Skin Injection – 1 Session + LITUO – 1 Vial"},"zh":{"e":"高端抗衰老干细胞水光针 + Re2O 1瓶"},"th":{"e":"ฉีด High-End Anti-Aging Stem Cell Skin RE2o 1 ขวด"},"vi":{"e":"Tiêm skin tế bào gốc chống lão hóa high-end 1 lần + re2o 1 lọ"}}},{"k":"미라셀 줄기세포 프로그램","c":"줄기세포","t":{"en":{"e":"Miracell Stem Cell Program"},"zh":{"e":"Miracell 干细胞项目"},"th":{"e":"Miracell Stem Cell Program"},"vi":{"e":"Liệu trình tế bào gốc Miracell"}}},{"k":"울쎄라 더블리프팅(턱밑울쎄라100샷+턱라인 민트파인4줄)","c":"리프팅","t":{"en":{"e":"Ultherapy Double Lifting (100 Shots Under-Chin Ultherapy + 4 Mint Fine Threads for Jawline)"},"zh":{"e":"Ulthera 双重提升（下巴 Ulthera 100发 + 下颌线 Mint Fine 4条）"},"th":{"e":"อัลเทอร่า Double Lifting (อัลเทอร่าใต้คาง 100 ช็อต + Mint Fine กรอบหน้า 4 เส้น)"},"vi":{"e":"Ultherapy double lifting (Ulthera 100 shot vùng dưới cằm + chỉ Mint Fine 4 sợi vùng đường viền hàm)"}}},{"k":"울쎄라 300샷","c":"리프팅","t":{"en":{"e":"Ultherapy 300 Shots"},"zh":{"e":"Ulthera 300发"},"th":{"e":"อัลเทอร่า 300 ช็อต"},"vi":{"e":"Ultherapy 300 shot"}}},{"k":"써마지 300샷","c":"리프팅","t":{"en":{"e":"Thermage 300 Shots"},"zh":{"e":"Thermage 300发"},"th":{"e":"เทอร์มาจ 300 ช็อต"},"vi":{"e":"Thermage 300 shot"},"ja":{"e":"サーマクール 300ショット"},"zh-TW":{"e":"第四代鳳凰電波 300發"}}},{"k":"덴서티 300샷","c":"리프팅","t":{"en":{"e":"DENSITY 300 Shots"},"zh":{"e":"Density 300发"},"th":{"e":"Density 300 ช็อต"},"vi":{"e":"Density 300 shot"}}},{"k":"덴서티 600샷","c":"리프팅","t":{"en":{"e":"DENSITY 600 Shots"},"zh":{"e":"Density 600发"},"th":{"e":"Density 600 ช็อต"},"vi":{"e":"Density 600 shot"}}},{"k":"슈링크 300샷","c":"리프팅","t":{"en":{"e":"Shurink 300 Shots"},"zh":{"e":"Shurink 300发"},"th":{"e":"Shurink 300 ช็อต"},"vi":{"e":"Shurink 300 shot"}}},{"k":"리니어지 300샷","c":"리프팅","t":{"en":{"e":"LinearZ 300 Shots"},"zh":{"e":"Lineage 300发"},"th":{"e":"LinearZ 300 ช็อต"},"vi":{"e":"Linear z 300 shot"}}},{"k":"인모드 fx+포마 1회","c":"리프팅","t":{"en":{"e":"InMode FX + FORMA – 1 Session"},"zh":{"e":"InMode FX+Forma 1次"},"th":{"e":"InMode fx+Forma 1 ครั้ง"},"vi":{"e":"InMode FX + Forma 1 lần"}}},{"k":"리투오 1병","c":"스킨부스터","t":{"en":{"e":"LITUO – 1 Vial"},"zh":{"e":"Re2O 1瓶"},"th":{"e":"RE2o 1 ขวด"},"vi":{"e":"re2o 1 lọ"}}},{"k":"리투오 2병","c":"스킨부스터","t":{"en":{"e":"LITUO – 2 Vials"},"zh":{"e":"Re2O 2瓶"},"th":{"e":"RE2o 2 ขวด"},"vi":{"e":"re2o 2 lọ"}}},{"k":"리투오 3병","c":"스킨부스터","t":{"en":{"e":"LITUO – 3 Vials"},"zh":{"e":"Re2O 3瓶"},"th":{"e":"RE2o 3 ขวด"},"vi":{"e":"re2o 3 lọ"}}},{"k":"리투오 1회+프리미엄토닝+비타민관리 10회","c":"스킨부스터","t":{"en":{"e":"LITUO 1 Session + Premium Toning + 10 Vitamin Care Sessions"},"zh":{"e":"Re2O 1次 + 高级嫩肤 + 维他命护理 10次"},"th":{"e":"RE2o 1 ครั้ง+Premium Toning+วิตามินผิว 10 ครั้ง"},"vi":{"e":"re2o 1 lần + điều trị toning premium + chăm sóc vitamin 10 lần"}}},{"k":"리투오 1병+쥬베룩1병 (스컬트라로 변경시 20만 추가)","c":"스킨부스터","t":{"en":{"e":"LITUO 1 Vial + 1 Vial Juvelook (Upgrade to Sculptra +KRW 200,000)"},"zh":{"e":"Re2O 1瓶 + Juvelook 1瓶（更换为 Sculptra 需另加20万韩元）"},"th":{"e":"RE2o 1 ขวด+Juvelook 1 ขวด (เปลี่ยนเป็น Sculptra เพิ่ม 2 แสนวอน)"},"vi":{"e":"re2o 1 lọ + Juvelook 1 lọ (nếu đổi sang Sculptra sẽ cộng thêm 200.000 KRW)"}}},{"k":"쥬베룩아이 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 1 Vial"},"zh":{"e":"Juvelook Eye 1瓶"},"th":{"e":"Juvelook Eye 1 ขวด"},"vi":{"e":"Juvelook Eye 1 lọ"}}},{"k":"쥬베룩아이 3병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 3 Vials"},"zh":{"e":"Juvelook Eye 3瓶"},"th":{"e":"3 ขวด"},"vi":{"e":"Juvelook Eye 3 lọ"}}},{"k":"쥬베룩아이1병+아이리쥬란1cc+아이스킨보톡스","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye 1 Vial + Eye Rejuran 1cc + Ice Skin Botox"},"zh":{"e":"Juvelook Eye1瓶 + 眼部 Rejuran 1cc + 眼部肉毒素"},"th":{"e":"1 ขวด + Eye Rejuran 1cc + สกินโบท็อกสำหรับตา"},"vi":{"e":"Juvelook Eye 1 lọ + Rejuran Eye 1cc + Botox skin Ice"}}},{"k":"마운자로1단계 2.5mg 1box(4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 1 – 2.5mg (1 Box / 4 Pens)"},"zh":{"e":"Mounjaro 第1阶段 2.5mg 1盒（4支）"},"th":{"e":"Mounjaro Step 1 2.5mg 1 กล่อง (4 pens)"},"vi":{"e":"Mounjaro giai đoạn 1 2.5mg 1 hộp (4 bút tiêm)"}}},{"k":"마운자로1단계 2.5mg 3box(12pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 1 – 2.5mg (3 Boxes / 12 Pens)"},"zh":{"e":"Mounjaro 2.5mg 3盒（12支）"},"th":{"e":"Mounjaro Step 1 2.5mg 3 กล่อง (12 pens)"},"vi":{"e":"Mounjaro giai đoạn 1 2,5mg 3 hộp (12 bút tiêm)"}}},{"k":"마운자로2단계 5.0mg 1box(4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 2 – 5.0mg (1 Box / 4 Pens)"},"zh":{"e":"Mounjaro 第2阶段 5.0mg 1盒（4支）"},"th":{"e":"Mounjaro Step 2 5.0mg 1 กล่อง (4 pens)"},"vi":{"e":"Mounjaro giai đoạn 2 5.0mg 1 hộp (4 bút tiêm)"}}},{"k":"마운자로2단계 5.0mg 3box(14pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 2 –5.0mg (3 Boxes / 14 Pens)"},"zh":{"e":"Mounjaro 5.0mg 3盒（14支）"},"th":{"e":"Mounjaro Step 2 5.0mg 3 กล่อง (14 pens)"},"vi":{"e":"Mounjaro giai đoạn 2 5,0mg 3 hộp (14 bút tiêm)"}}},{"k":"마운자로 3단계 7.5mg 1box (4pens)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro Step 3 – 7.5mg (1 Box / 4 Pens)"},"zh":{"e":"Mounjaro 第3阶段 7.5mg 1盒（4支）"},"th":{"e":"Mounjaro Step 3 7.5mg 1 กล่อง (4 pens)"},"vi":{"e":"Mounjaro giai đoạn 3 - 7,5mg 1 hộp (4 bút tiêm)"}}},{"k":"여자 겨드랑이 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal (Apogee) – 1 Session"},"zh":{"e":"女性腋下脱毛（Apogee）1次"},"th":{"e":"เลเซอร์รักแร้ผู้หญิง (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông nách nữ (Apogee) 1 lần"}}},{"k":"여자 겨드랑이 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"女性腋下脱毛（Apogee）5次"},"th":{"e":"เลเซอร์รักแร้ผู้หญิง (Apogee) 5 ครั้ง"},"vi":{"e":"Triệt lông nách nữ (Apogee) 5 lần"}}},{"k":"여자 팔 전체 제모(아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Full Arm Hair Removal (Apogee) – 1 Session"},"zh":{"e":"女性全臂脱毛（Apogee）1次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้หญิง (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ cánh tay nữ (Apogee) 1 lần"}}},{"k":"여자 팔 전체 제모(아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Full Arm Hair Removal (Apogee) –  5 Sessions"},"zh":{"e":"女性全臂脱毛（Apogee） 5次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้หญิง 5 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ cánh tay nữ (Apogee)5 lần"}}},{"k":"여자 다리 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Women’s Full Leg Hair Removal (Apogee) – 1 Session"},"zh":{"e":"女性全腿脱毛（Apogee）1次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้หญิง (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ chân nữ (Apogee) 1 lần"}}},{"k":"여자 다리 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Women’s Full Leg Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"女性全腿脱毛（Apogee）5次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้หญิง 5 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ chân nữ (Apogee) 5 lần"}}},{"k":"남자 겨드랑이 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal (Apogee) – 1 Session"},"zh":{"e":"男性腋下脱毛（Apogee）1次"},"th":{"e":"เลเซอร์รักแร้ผู้ชาย (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông nách nam (Apogee) 1 lần"}}},{"k":"남자 겨드랑이 제모(아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"男性腋下脱毛（Apogee）5次"},"th":{"e":"เลเซอร์รักแร้ผู้ชาย 5 ครั้ง"},"vi":{"e":"Triệt lông nách nam (Apogee) 5 lần"}}},{"k":"남자 팔 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Arm Hair Removal (Apogee) – 1 Session"},"zh":{"e":"男性全臂脱毛（Apogee）1次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้ชาย (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ cánh tay nam (Apogee) 1 lần"}}},{"k":"남자 팔 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Arm Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"男性全臂脱毛（Apogee）5次"},"th":{"e":"เลเซอร์แขนทั้งแขนผู้ชาย 5 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ cánh tay nam (Apogee) 5 lần"}}},{"k":"남자 다리 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Leg Hair Removal (Apogee) – 1 Session"},"zh":{"e":"男性全腿脱毛（Apogee）1次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้ชาย (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ chân nam (Apogee) 1 lần"}}},{"k":"남자 다리 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Leg Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"男性全腿脱毛（Apogee）5次"},"th":{"e":"เลเซอร์ขาทั้งขาผู้ชาย 5 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ chân nam (Apogee) 5 lần"}}},{"k":"남자 수염 전체 제모 (아포지) 1회","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal (Apogee) – 1 Session"},"zh":{"e":"男性全胡须脱毛（Apogee）1次"},"th":{"e":"เลเซอร์หนวดทั่วหน้าผู้ชาย (Apogee) 1 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ râu nam (Apogee) 1 lần"}}},{"k":"남자 수염 전체 제모 (아포지) 5회","c":"제모","t":{"en":{"e":"Men’s Full Beard Hair Removal (Apogee) – 5 Sessions"},"zh":{"e":"男性全胡须脱毛（Apogee）5次"},"th":{"e":"เลเซอร์หนวดทั่วหน้าผู้ชาย 5 ครั้ง"},"vi":{"e":"Triệt lông toàn bộ râu nam (Apogee) 5 lần"}}},{"k":"큐오필 4cc","c":"필러","t":{"en":{"e":"QO Fill 4cc"},"zh":{"e":"QO Fill 4cc"},"th":{"e":"QO Fill 4cc"},"vi":{"e":"Q-O Fill 4cc"}}},{"k":"큐오필 8cc","c":"필러","t":{"en":{"e":"QO Fill 8cc"},"zh":{"e":"QO Fill 8cc"},"th":{"e":"QO Fill 8cc"},"vi":{"e":"Q-O Fill 8cc"}}},{"k":"큐오필 12cc","c":"필러","t":{"en":{"e":"QO Fill 12cc"},"zh":{"e":"QO Fill 12cc"},"th":{"e":"QO Fill 12cc"},"vi":{"e":"Q-O Fill 12cc"}}},{"k":"레디어스 1병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 1 Vial"},"zh":{"e":"Radiesse 1瓶"},"th":{"e":"Radiesse 1 ขวด"},"vi":{"e":"Radiesse 1 lọ"}}},{"k":"레디어스 2병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 2 Vials"},"zh":{"e":"Radiesse 2瓶"},"th":{"e":"Radiesse 2 ขวด"},"vi":{"e":"Radiesse 2 lọ"}}},{"k":"레디어스 3병","c":"스킨부스터","t":{"en":{"e":"Radiesse – 3 Vials"},"zh":{"e":"Radiesse 3瓶"},"th":{"e":"Radiesse 3 ขวด"},"vi":{"e":"Radiesse 3 lọ"}}},{"k":"피코프락셀 (풀페이스) 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel (Full Face) – 1 Session"},"zh":{"e":"Pico Fraxel（全脸）1次"},"th":{"e":"พิโคแฟรคเซล ทั่วหน้า 1 ครั้ง"},"vi":{"e":"Pico Fraxel (toàn mặt) 1 lần"}}},{"k":"미라젯+쥬베룩(풀페이스)1회","c":"여드름/모공/흉터","t":{"en":{"e":"MiraJet + Juvelook (Full Face) – 1 Session"},"zh":{"e":"Mirajet + Juvelook（全脸）1次"},"th":{"e":"Mirajet + Juvelook ทั่วหน้า 1 ครั้ง"},"vi":{"e":"Mirajet + Juvelook (toàn mặt) 1 lần"}}},{"k":"포텐자(풀페이스)1회","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza (Full Face) – 1 Session"},"zh":{"e":"Potenza（全脸）1次"},"th":{"e":"โพเทนซ่า ทั่วหน้า 1 ครั้ง"},"vi":{"e":"Potenza (toàn mặt) 1 lần"}}},{"k":"토닝5회+클라리티1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"5 Toning Sessions + 1 Clarity Session"},"zh":{"e":"嫩肤 5次 + Clarity 1次"},"th":{"e":"โทนนิ่ง 5 ครั้ง + Clarity 1 ครั้ง"},"vi":{"e":"Toning 5 lần + Clarity 1 lần"}}},{"k":"리투오1병+화이트닝토닝10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"LITUO 1 Vial + 10 Whitening Toning Sessions"},"zh":{"e":"Re2O 1瓶 + 美白嫩肤10次"},"th":{"e":"RE2o 1 ขวด + โทนนิ่งผิวขาว 10 ครั้ง"},"vi":{"e":"re2o 1 lọ + toning làm trắng 10 lần"}}},{"k":"리투오1병+프리미엄토닝+비타민관리10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"LITUO 1 Vial + Premium Toning + 10 Vitamin Care Sessions"},"zh":{"e":"Re2O 1瓶 + 高级嫩肤 + 维他命护理10次"},"th":{"e":"RE2o 1 ขวด + พรีเมียมโทนนิ่ง + วิตามินทรีตเมนต์ 10 ครั้ง"},"vi":{"e":"re2o 1 lọ + toning premium + chăm sóc vitamin 10 lần"}}},{"k":"리쥬란힐러 2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran Healer 2cc"},"zh":{"e":"Rejuran Healer 2cc"},"th":{"e":"Rejuran Healer 2cc"},"vi":{"e":"Rejuran Healer 2cc"}}},{"k":"리쥬란hb  2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran HB 2cc"},"zh":{"e":"Rejuran hb 2cc"},"th":{"e":"Rejuran HB 2cc"},"vi":{"e":"Rejuran HB 2cc"}}},{"k":"쥬베룩 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook – 1 Vial"},"zh":{"e":"Juvelook 1瓶"},"th":{"e":"Juvelook 1 ขวด"},"vi":{"e":"Juvelook 1 lọ"}}},{"k":"온다5만줄+1만줄","c":"리프팅","t":{"en":{"e":"ONDA 50,000 Shots + 10,000 Bonus Shots"},"zh":{"e":"ONDA 5万发 + 1万发"},"th":{"e":"Onda 50KJ + 10KJ"},"vi":{"e":"Onda 50.000 sợi + 10.000 sợi"}}},{"k":"온다10만줄+2만줄","c":"리프팅","t":{"en":{"e":"ONDA 100,000 Shots + 20,000 Bonus Shots"},"zh":{"e":"ONDA 10万发 + 2万发"},"th":{"e":"Onda 100KJ + 20KJ"},"vi":{"e":"Onda 100.000 sợi + 20.000 sợi"}}},{"k":"민트 파인 4가닥","c":"실리프팅","t":{"en":{"e":"Mint Fine Threads – 4 Threads"},"zh":{"e":"Mint Fine 4根"},"th":{"e":"Mint Fine 4 เส้น"},"vi":{"e":"Mint Fine 4 sợi"}}},{"k":"잼버 4가닥","c":"실리프팅","t":{"en":{"e":"Jamber Threads – 4 Threads"},"zh":{"e":"Jamber 4根"},"th":{"e":"Jamber 4 เส้น"},"vi":{"e":"Jamber 4 sợi"}}},{"k":"팽팽 민트 모노 100가닥","c":"실리프팅","t":{"en":{"e":"Tightening Mint Mono Threads – 100 Threads"},"zh":{"e":"紧致 Mint Mono 100根"},"th":{"e":"Mint Mono 100 เส้น"},"vi":{"e":"Chỉ Mint Mono căng 100 sợi"}}},{"k":"이마폭스 실리프팅","c":"실리프팅","t":{"en":{"e":"Forehead Fox Eye Thread Lift"},"zh":{"e":"额头 Fox 线雕"},"th":{"e":"ร้อยไหม Foxy Eyes หน้าผาก"},"vi":{"e":"Căng chỉ Fox vùng trán"}}},{"k":"내맘대로 커스텀 실리프팅(미미썸 시그니처)","c":"실리프팅","t":{"en":{"e":"Customized Thread Lift (MiMiSome Signature)"},"zh":{"e":"自由定制线雕（Mimisome Signature）"},"th":{"e":"ร้อยไหม Custom Design (Mimisome Signature)"},"vi":{"e":"Căng chỉ custom theo yêu cầu (MimiSome signature)"}}},{"k":"골반필러 100cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 100cc"},"zh":{"e":"骨盆填充 100cc"},"th":{"e":"ฟิลเลอร์สะโพก 100cc"},"vi":{"e":"Filler vùng hông  100cc"}}},{"k":"골반필러 200cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 200cc"},"zh":{"e":"骨盆填充 200cc"},"th":{"e":"ฟิลเลอร์สะโพก 200cc"},"vi":{"e":"Filler vùng hông  200cc"}}},{"k":"골반필러 300cc","c":"필러","t":{"en":{"e":"Hip Dip / Pelvic Filler 300cc"},"zh":{"e":"骨盆填充 300cc"},"th":{"e":"ฟิลเลอร์สะโพก 300cc"},"vi":{"e":"Filler vùng hông  300cc"}}},{"k":"가슴볼륨필러 100cc","c":"필러","t":{"en":{"e":"Breast Volume Filler 100cc"},"zh":{"e":"胸部丰盈填充 100cc"},"th":{"e":"ฟิลเลอร์หน้าอก 100cc"},"vi":{"e":"Filler tăng thể tích ngực 100cc"}}},{"k":"직각어깨필러","c":"필러","t":{"en":{"e":"Right-Angle Shoulder Filler"},"zh":{"e":"直角肩填充"},"th":{"e":"ฟิลเลอร์ไหล่ตรง"},"vi":{"e":"Filler tạo vai vuông"}}},{"k":"제니직각어깨필러 20cc+승모근보톡스200유닛(필러 추가시 1cc 59900원)","c":"필러","t":{"en":{"e":"Jenny Right-Angle Shoulder Filler 20cc + 200 Units+Trapezius Botox (Additional Filler KRW 59,900/cc)"},"zh":{"e":"Jennie 直角肩填充 20cc + 斜方肌肉毒素 200单位（追加填充每1cc为59900韩元）"},"th":{"e":"Jenny ฟิลเลอร์ไหล่ตรง 20cc + โบท็อกซ์บ่า 200 ยูนิต(เพิ่มฟิลเลอร์ 1cc = 59,900 วอน)"},"vi":{"e":"Filler vai vuông kiểu vai Jennie 20cc + Botox cơ thang 200 units (nếu thêm filler: 59.900 KRW / 1cc)"}}},{"k":"미니 spl 한달무제한 (인바디+비만약처방전포함)","c":"비만/다이어트","t":{"en":{"e":"Mini SPL Unlimited Monthly Package (Includes InBody Test + Obesity Medication Prescription)"},"zh":{"e":"迷你 SPL 一个月不限次数（含 InBody 检测 + 减肥药处方）"},"th":{"e":"Mini SPL ไม่จำกัด 1 เดือน (รวม Inbody + ใบสั่งยาลดน้ำหนัก)"},"vi":{"e":"Mini SPL không giới hạn 1 tháng (bao gồm đo InBody + đơn thuốc giảm cân)"}}},{"k":"지방 추출 주사 S펫 1병","c":"비만/다이어트","t":{"en":{"e":"S-FAT Fat Extraction Injection – 1 Vial"},"zh":{"e":"抽脂针 S Pet 1瓶"},"th":{"e":"S-Fat 1 ขวด"},"vi":{"e":"Tiêm chiết xuất mỡ S-Fat 1 lọ"}}},{"k":"지방분해주사 한달무제한","c":"비만/다이어트","t":{"en":{"e":"Unlimited Fat-Dissolving Injections – 1 Month"},"zh":{"e":"溶脂针 一个月不限次数"},"th":{"e":"ฉีดสลายไขมัน ไม่จำกัด 1 เดือน"},"vi":{"e":"Tiêm tan mỡ không giới hạn 1 tháng"}}},{"k":"위고비 1개월+바디온다 1개월 (8만줄)","c":"비만/다이어트","t":{"en":{"e":"Wegovy 1-Month Program + Body ONDA 1 Month (80,000 Shots)"},"zh":{"e":"Wegovy 1个月 + Body ONDA 1个月（8万发）"},"th":{"e":"Wegovy 1 เดือน + Body Onda 1 เดือน (80KJ)"},"vi":{"e":"Wegovy 1 tháng + Body Onda 1 tháng (80.000 sợi)"}}},{"k":"마운자로 1개월 +바디온다 1개월(8만줄)","c":"비만/다이어트","t":{"en":{"e":"Mounjaro 1-Month Program + Body ONDA 1 Month (80,000 Shots)"},"zh":{"e":"Mounjaro 1个月 + Body ONDA 1个月（8万发）"},"th":{"e":"Mounjaro 1 เดือน + Body Onda 1 เดือน (80KJ)"},"vi":{"e":"Mounjaro 1 tháng + Body Onda 1 tháng (80.000 sợi)"}}},{"k":"얼굴 점제거 1개당","kd":"0.2mm 이하 크기, 10개까지 / 첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Mole Removal – Per Lesion","d":"Up to 10 lesions under 0.2mm / First Visit Trial Price"},"vi":{"e":"Xóa nốt ruồi trên mặt 1 nốt","d":"Kích thước dưới 0,2mm, tối đa 10 cái / giá trải nghiệm lần đầu đến"}}},{"k":"흑자, 검버섯 1개","kd":"크기별 상이, 0.5cm 이내 / 첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Age Spot / Solar Lentigo Removal – Per Lesion","d":"Price Varies by Size, Within 0.5cm / First Visit Trial Price"},"vi":{"e":"Đốm sắc tố, đốm lão hóa 1 cái","d":"Giá khác nhau theo kích thước, trong phạm vi 0,5cm / giá trải nghiệm lần đầu đến"}}},{"k":"편평사마귀 / 쥐젖 / 비립종 1개당","kd":"첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Flat Wart / Skin Tag / Milia Removal – Per Lesion","d":"First Visit Trial Price"},"vi":{"e":"Mụn cóc phẳng, skin tag, hạt kê 1 cái","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"아쿠아필 1제, 2제","kd":"첫방문 체험가","c":"여드름/모공/흉터","t":{"en":{"e":"Aqua Peel Step 1 & 2","d":"First Visit Trial Price"},"vi":{"e":"Aquapeel bước 1, bước 2","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"주름 보톡스 (국산)","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"Wrinkle Botox (Korean Brand)","d":"First Visit Trial Price"},"vi":{"e":"Botox xóa nhăn (nội địa)","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"사각턱 보톡스 (국산)","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox (Korean Brand)","d":"First Visit Trial Price"},"vi":{"e":"Botox gọn hàm vuông (nội địa)","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"브이핏 윤곽주사 1cc","kd":"첫방문 체험가","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection 1cc","d":"First Visit Trial Price"},"vi":{"e":"Tiêm tạo V-line 1cc","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"프리미엄 명품 피코토닝 1회","kd":"첫방문 체험가","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Premium Pico Toning – 1 Session","d":"First Visit Trial Price"},"vi":{"e":"Pico toning premium cao cấp 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"여성 아포지 인중 or 겨드랑이 제모 1회","kd":"첫방문 체험가","c":"제모","t":{"en":{"e":"Women’s Apogee Upper Lip or Underarm Hair Removal – 1 Session","d":"First Visit Trial Price"},"vi":{"e":"Triệt lông Apogee nữ vùng ria mép hoặc nách 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"남성 아포지 인중 제모 1회","kd":"첫방문 체험가","c":"제모","t":{"en":{"e":"Men’s Apogee Upper Lip Hair Removal – 1 Session","d":"First Visit Trial Price"},"vi":{"e":"Triệt lông Apogee nam vùng ria mép 1 lần","d":"Giá trải nghiệm lần đầu đến"}}},{"k":"표정근육주름 1부위","kd":"하이톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Hutox (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vi":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"HiTox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"코어톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Coretox  (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vi":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"CoreTox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"이노톡스 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Innotox  (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vi":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"Innotox (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"표정근육주름 1부위","kd":"제오민 (이마, 미간, 눈가, 콧대, 콧등, 자갈 중 선택)","c":"보톡스/윤곽주사","t":{"en":{"e":"Dynamic Wrinkle Botox – 1 Area","d":"Xeomin  (Choose 1: Forehead / Frown Lines / Crow’s Feet / Nasal Bridge / Bunny Lines / Pebble Chin)"},"vi":{"e":"Nếp nhăn cơ biểu cảm 1 vùng","d":"Xeomin (chọn 1 trong: trán, giữa chân mày, đuôi mắt, sống mũi, lưng mũi, cằm đá cuội)"}}},{"k":"사각턱 보톡스","kd":"하이톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Hutox"},"vi":{"e":"Botox gọn cằm vuông","d":"HiTox"}}},{"k":"사각턱 보톡스","kd":"코어톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Coretox"},"vi":{"e":"Botox gọn cằm vuông","d":"CoreTox"}}},{"k":"사각턱 보톡스","kd":"이노톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Innotox"},"vi":{"e":"Botox gọn cằm vuông","d":"Innotox"}}},{"k":"사각턱 보톡스","kd":"제오민","c":"보톡스/윤곽주사","t":{"en":{"e":"Jaw Botox / Masseter Botox","d":"Xeomin"},"vi":{"e":"Botox gọn cằm vuông","d":"Xeomin"}}},{"k":"스킨보톡스","kd":"하이톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Hutox"},"vi":{"e":"Skin Botox","d":"HiTox"}}},{"k":"스킨보톡스","kd":"코어톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Coretox"},"vi":{"e":"Skin Botox","d":"CoreTox"}}},{"k":"스킨보톡스","kd":"이노톡스","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Innotox"},"vi":{"e":"Skin Botox","d":"Innotox"}}},{"k":"스킨보톡스","kd":"제오민","c":"보톡스/윤곽주사","t":{"en":{"e":"Skin Botox","d":"Xeomin"},"vi":{"e":"Skin Botox","d":"Xeomin"}}},{"k":"브이핏 윤곽주사 12cc","c":"보톡스/윤곽주사","t":{"en":{"e":"V-FIT Contouring Injection 12cc"},"vi":{"e":"Tiêm tạo V-line 12cc"}}},{"k":"미미썸 양악주사","c":"보톡스/윤곽주사","t":{"en":{"e":"MiMiSome Double Jaw Contouring Injection"},"vi":{"e":"Tiêm tạo gọn hàm Mimisome"}}},{"k":"코 조각주사","c":"보톡스/윤곽주사","t":{"en":{"e":"Nose Contouring Injection"},"vi":{"e":"Tiêm tạo dáng mũi"}}},{"k":"볼륨 필러 1cc","kd":"아띠에르","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Attiére"},"vi":{"e":"Filler tạo volume 1cc","d":"Atiere"}}},{"k":"볼륨 필러 1cc","kd":"아말리안","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Amalian"},"vi":{"e":"Filler tạo volume 1cc","d":"Amalian"}}},{"k":"볼륨 필러 1cc","kd":"쥬비덤","c":"필러","t":{"en":{"e":"Volume Filler 1cc","d":"Juvederm"},"vi":{"e":"Filler tạo volume 1cc","d":"Juvederm"}}},{"k":"벨로테로 소프트 1cc","kd":"눈두덩 & 눈밑고랑 꺼짐 개선","c":"필러","t":{"en":{"e":"Belotero Soft 1cc","d":"Upper Eyelid & Tear Trough Volume Correction"},"vi":{"e":"Belotero Soft 1cc","d":"Cải thiện vùng mí mắt & rãnh dưới mắt bị hõm"}}},{"k":"레스틸렌 키스 1cc","c":"필러","t":{"en":{"e":"Restylane Kysse 1cc"},"vi":{"e":"Restylane Kysse 1cc"}}},{"k":"큐오필 줄기세포","c":"필러","t":{"en":{"e":"QO Fill Stem Cell Filler"},"vi":{"e":"Q-O Fill tế bào gốc"}}},{"k":"인모드 FX 1부위","c":"리프팅","t":{"en":{"e":"InMode FX – 1 Area"},"vi":{"e":"InMode FX 1 vùng"}}},{"k":"인모드 FORMA","c":"리프팅","t":{"en":{"e":"InMode FORMA"},"vi":{"e":"InMode Forma"}}},{"k":"온다 3만줄 + LDM 초음파 리프팅 12분","c":"리프팅","t":{"en":{"e":"ONDA 30,000 Shots + LDM Ultrasound Lifting 12 min"},"vi":{"e":"Onda 30.000 lines + LDM nâng cơ bằng sóng siêu âm 12 phút"}}},{"k":"물광주사 6cc","c":"스킨부스터","t":{"en":{"e":"Skin Booster Injection 6cc"},"vi":{"e":"Tiêm căng bóng 6cc"}}},{"k":"리쥬란 2cc","c":"스킨부스터","t":{"en":{"e":"Rejuran 2cc"},"vi":{"e":"Rejuran 2cc"}}},{"k":"쥬베룩 스킨 4cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Skin 4cc"},"vi":{"e":"Juvelook Skin 4cc"}}},{"k":"쥬베룩 볼륨 2cc","c":"스킨부스터","t":{"en":{"e":"Juvelook Volume 2cc"},"vi":{"e":"Juvelook Volume 2cc"}}},{"k":"쥬베룩 아이 1병","c":"스킨부스터","t":{"en":{"e":"Juvelook Eye – 1 Vial"},"vi":{"e":"Juvelook Eye 1 lọ"}}},{"k":"미라셀 스킨주사 30cc","c":"줄기세포","t":{"en":{"e":"Miracell Skin Injection 30cc"},"vi":{"e":"Tiêm skin Miracell 30cc"}}},{"k":"미라셀 정맥주사 120cc","c":"줄기세포","t":{"en":{"e":"Miracell  IV Therapy 120cc"},"vi":{"e":"Tiêm Truyền tĩnh mạch 120cc"}}},{"k":"민트실 1줄","c":"실리프팅","t":{"en":{"e":"Mint Thread – 1 Thread"},"vi":{"e":"Chỉ Mint 1 sợi"}}},{"k":"민트실 4줄","c":"실리프팅","t":{"en":{"e":"Mint Thread – 4 Threads"},"vi":{"e":"Chỉ Min 4 sợi"}}},{"k":"눈썹 / 이마 거상 실리프팅 4줄","c":"실리프팅","t":{"en":{"e":"Brow / Forehead Lift Thread Lifting – 4 Threads"},"vi":{"e":"Căng chỉ mày / trán  4 sợi"}}},{"k":"하이코 실리프팅 2줄","c":"실리프팅","t":{"en":{"e":"HICO Nose Thread Lift – 2 Threads"},"vi":{"e":"Căng chỉ High-Ko 2 sợi"}}},{"k":"브라이톤 토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Brightone Toning – 1 Session"},"vi":{"e":"Brighton Toning  1 lần"}}},{"k":"포텐자 (모공 / 기미 / 여드름) + 연어 재생 앰플 관리 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Potenza (Pores / Melasma / Acne) + Salmon Regeneration Ampoule Care – 1 Session"},"vi":{"e":"Potenza (lỗ chân lông / nám / mụn) + chăm sóc ampoule tái tạo từ cá hồi 1 lần"}}},{"k":"엑셀브이 플러스 제네시스 + 진정 앰플 관리 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Excel V Plus Genesis + Soothing Ampoule Care – 1 Session"},"vi":{"e":"Excel V Plus Genesis + chăm sóc ampoule làm dịu da 1 lần"}}},{"k":"헐리우드 스펙트라 프리미엄 토닝 1회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Premium Toning – 1 Session"},"vi":{"e":"Hollywood Spectra premium toning 1 lần"}}},{"k":"주근깨, 잡티 레이저 1회","kd":"클라리티 Pro","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Freckle & Pigmentation Laser – 1 Session","d":"Pro Clarity Pro"},"vi":{"e":"Laser trị tàn nhang, đốm sắc tố 1 lần","d":"Clarity Pro"}}},{"k":"브라이톤 듀얼토닝 10회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Brightone Dual Toning – 10 Sessions"},"vi":{"e":"Brighton dual toning 10 lần"}}},{"k":"헐리우드 스펙트라 프리미엄 토닝 12회","c":"피부레이저(색소/혈관/기미)","t":{"en":{"e":"Hollywood Spectra Premium Toning – 12 Sessions"},"vi":{"e":"Hollywood Spectra premium toning 12 lần"}}},{"k":"아쿠아필 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Aqua Peel – 1 Session"},"vi":{"e":"Aquapeel 1 lần"}}},{"k":"라라필 1회","c":"여드름/모공/흉터","t":{"en":{"e":"LHALALA Peel – 1 Session"},"vi":{"e":"Lhala Peel 1 lần"}}},{"k":"LDM 재생관리 1회","c":"여드름/모공/흉터","t":{"en":{"e":"LDM Regeneration Care – 1 Session"},"vi":{"e":"Chăm sóc tái tạo LDM 1 lần"}}},{"k":"여드름관리 1회","c":"여드름/모공/흉터","t":{"en":{"e":"Acne Care Treatment – 1 Session"},"vi":{"e":"Chăm sóc điều trị mụn 1 lần"}}},{"k":"피코프락셀 1회","kd":"부분","c":"여드름/모공/흉터","t":{"en":{"e":"Pico Fraxel – 1 Session","d":"Partial Area"},"vi":{"e":"Pico Fraxel 1 lần","d":"1 phần"}}},{"k":"여자 인중 제모 5회","c":"제모","t":{"en":{"e":"Women’s Upper Lip Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông ria mép nữ 5 lần"}}},{"k":"여자 겨드랑이 제모 5회","c":"제모","t":{"en":{"e":"Women’s Underarm Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông nách nữ  5 lần"}}},{"k":"여자 비키니라인 제모 5회","c":"제모","t":{"en":{"e":"Women’s Bikini Line Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông vùng bikini nữ  5 lần"}}},{"k":"여자 브라질리언 제모 5회","c":"제모","t":{"en":{"e":"Women’s Brazilian Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông toàn vùng kín nữ  5 lần"}}},{"k":"남자 인중 제모 5회","c":"제모","t":{"en":{"e":"Men’s Upper Lip Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông ria mép nam 5 lần"}}},{"k":"남자 하관전체 제모 5회","c":"제모","t":{"en":{"e":"Men’s Full Lower Face Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông toàn bộ vùng hàm dưới nam 5 lần"}}},{"k":"남자 겨드랑이 제모 5회","c":"제모","t":{"en":{"e":"Men’s Underarm Hair Removal – 5 Sessions"},"vi":{"e":"Triệt lông nách nam 5 lần"}}},{"k":"얼굴 전체 솜털 제모 1회","c":"제모","t":{"en":{"e":"Full Face Fine Hair Removal – 1 Session"},"vi":{"e":"Triệt lông tơ toàn mặt 1 lần"}}},{"k":"퀵영양수액 1회","kd":"태반 / 비타민/ 백옥 / 신데렐라 중 선택","c":"비만/다이어트","t":{"en":{"e":"Quick Nutrition IV Drip – 1 Session","d":"Choose 1: Placenta / Vitamin / Glutathione / Cinderella IV"},"vi":{"e":"Truyền dinh dưỡng nhanh 1 lần"}}},{"k":"퀵영양수액 5회","kd":"태반 / 비타민/ 백옥 / 신데렐라 중 선택","c":"비만/다이어트","t":{"en":{"e":"Quick Nutrition IV Drip – 5 Sessions"},"vi":{"e":"Truyền dinh dưỡng nhanh 5 lần","d":"Chọn 1 trong: nhau thai / vitamin / Bạch Ngọc / Cinderella"}}},{"k":"고농도 메가비타민 1회","kd":"고농도 비타민C","c":"비만/다이어트","t":{"en":{"e":"High-Dose Mega Vitamin IV – 1 Session","d":"High-Dose Vitamin C"},"vi":{"e":"Mega vitamin nồng độ cao 1 lần"}}},{"k":"고농도 메가비타민 5회","kd":"고농도 비타민C","c":"비만/다이어트","t":{"en":{"e":"High-Dose Mega Vitamin IV – 5 Sessions"},"vi":{"e":"Mega vitamin nồng độ cao 5 lần","d":"Vitamin C nồng độ cao"}}},{"k":"항노화주사 1회","kd":"포도당 + 비타민C + 비타민B","c":"비만/다이어트","t":{"en":{"e":"Anti-Aging Injection – 1 Session","d":"Glucose + Vitamin C + Vitamin B"},"vi":{"e":"Tiêm chống lão hóa 1 lần"}}},{"k":"항노화주사 5회","kd":"포도당 + 비타민C + 비타민B","c":"비만/다이어트","t":{"en":{"e":"Anti-Aging Injection – 5 Sessions"},"vi":{"e":"Tiêm chống lão hóa 5 lần"}}},{"k":"신비주사 1회","kd":"포도당 + 비타민C + 신데렐라","c":"비만/다이어트","t":{"en":{"e":"Shinbi IV Drip – 1 Session"},"vi":{"e":"Tiêm Shinbi 1 lần"}}},{"k":"신비주사 5회","kd":"포도당 + 비타민C + 신데렐라","c":"비만/다이어트","t":{"en":{"e":"Shinbi IV Drip – 5 Sessions"},"vi":{"e":"Tiêm Shinbi 5 lần","d":"Glucose + vitamin C + vitamin B"}}},{"k":"오로라주사 1회","kd":"포도당 + 비타민C + 백옥","c":"비만/다이어트","t":{"en":{"e":"Aurora IV Drip – 1 Session","d":"Glucose + Vitamin C + Glutathione"},"vi":{"e":"Tiêm Aurora 1 lần"}}},{"k":"오로라주사 5회","kd":"포도당 + 비타민C + 백옥","c":"비만/다이어트","t":{"en":{"e":"Aurora IV Drip – 5 Sessions"},"vi":{"e":"Tiêm Aurora 5 lần","d":"Glucose + vitamin C + Bạch Ngọc"}}},{"k":"소녀주사 1회","kd":"고농도 비타민C + 신데렐라 + 백옥 + 태반","c":"비만/다이어트","t":{"en":{"e":"Girl IV Drip – 1 Session","d":"High-Dose Vitamin C + Cinderella + Glutathione + Placenta"},"vi":{"e":"Tiêm tiêm trẻ hóa 1 lần"}}},{"k":"소녀주사 5회","kd":"고농도 비타민C + 신데렐라 + 백옥 + 태반","c":"비만/다이어트","t":{"en":{"e":"Girl IV Drip – 5 Sessions"},"vi":{"e":"Tiêm tiêm trẻ hóa 5 lần","d":"Vitamin C nồng độ cao + Cinderella + Bạch Ngọc + nhau thai"}}},{"k":"와인주사 추가","kd":"모든 수액에 추가 가능","c":"비만/다이어트","t":{"en":{"e":"Wine Injection Add-On","d":"Can Be Added to Any IV Drip"},"vi":{"e":"Thêm tiêm Wine injection","d":"Có thể thêm vào tất cả loại truyền dịch"}}},{"k":"무통증 수면 마취","kd":"장시간 풀수면 마취","c":"성형수술","t":{"en":{"e":"Painless Sedation Anesthesia","d":"Long-Duration Full Sedation Anesthesia"},"vi":{"e":"Gây mê ngủ không đau","d":"Gây mê ngủ sâu thời gian dài"}}},{"k":"저통증 가스 마취","kd":"에어녹스 스마일 가스 마취","c":"성형수술","t":{"en":{"e":"Low-Pain Gas Anesthesia","d":"AirNox Smile Gas Sedation"},"vi":{"e":"Gây mê khí giảm đau mức thấp","d":"Gây mê khí AirNOx Smile gas"}}},{"k":"보조개 한쪽","kd":"긴보조개 디자인 가능","c":"성형수술","t":{"en":{"e":"Dimple Creation – One Side","d":"Long Dimple Design Available"},"vi":{"e":"Tạo lúm đồng tiền 1 bên","d":"Có thể thiết kế lúm đồng tiền dài"}}},{"k":"3D 눈밑 돌출 지방 제거 + 눈밑 & 앞볼 지방 이식","c":"성형수술","t":{"en":{"e":"3D Under-Eye Fat Removal + Under-Eye & Front Cheek Fat Grafting"},"vi":{"e":"Loại bỏ mỡ phồng dưới mắt 3D + cấy mỡ dưới mắt & vùng má trước"}}},{"k":"360도 무제한 지방 흡입","kd":"팔전체 / 상 or 하복부 / 허벅지 중 선택","c":"성형수술","t":{"en":{"e":"360° Unlimited Liposuction","d":"Choose 1: Full Arms / Upper or Lower Abdomen / Thighs"},"vi":{"e":"Hút mỡ 360 độ không giới hạn","d":"Chọn: toàn bộ cánh tay / bụng trên hoặc bụng dưới / đùi"}}},{"k":"미미썸 한정","t":{"en":{"e":"Mimisome Limited"},"ja":{"e":"ミミソム限定"},"zh-TW":{"e":"MIMISOME 限定"}}},{"k":"메조주사 + LED","t":{"en":{"e":"Mesotherapy Injection + LED"},"ja":{"e":"メソセラピー注射 + LED"},"zh-TW":{"e":"頭皮中胚層療法注射 + LED + 藥物處方"}}},{"k":"메조주사 + LED + 약물처방","t":{"en":{"e":"Mesotherapy Injection + LED + Medication Prescription"},"ja":{"e":"メソセラピー注射 + LED + 薬処方"},"zh-TW":{"e":"頭皮中胚層療法注射 + LED + 藥物處方"}}},{"k":"*VAT포함","t":{"en":{"e":"VAT included"},"ja":{"e":"VAT込み"},"zh-TW":{"e":"含增值稅"}}},{"k":"휑한 가르마 꽉! 잡는","t":{"en":{"e":"Targets thinning hair part lines"},"ja":{"e":"分け目の薄さをしっかりカバーする"},"zh-TW":{"e":"稀疏分線全面改善"}}},{"k":"3-Step 메디컬 탈모·두피 재생 솔루션","t":{"en":{"e":"A 3 step medical hair loss and scalp regeneration solution"},"ja":{"e":"3-Step 医療発毛・頭皮再生ソリューション"},"zh-TW":{"e":"3-Step 醫療級落髮頭皮再生方案"}}},{"k":"주사하고, 빛으로 채우는 빈틈없는 모근 강화 프로그램","t":{"en":{"e":"Fill from within by taking medication, injecting, and using light. Pore strengthening program"},"ja":{"e":"注入・光で補う隙のない毛根強化プログラム"},"zh-TW":{"e":"透過口服、注射,再以光能填補空隙的完整毛根強化計畫"}}},{"k":"이런 분들께 강력 추천합니다!","t":{"en":{"e":"Recommended for these people"},"ja":{"e":"次のような方に強くお勧めします!"},"zh-TW":{"e":"強烈推薦給以下族群!"}}},{"k":"하루가 다르게 머리카락이 얇아지고 힘이 없는 분","t":{"en":{"e":"Those whose hair part is widening day by day and showing no improvement"},"ja":{"e":"日々髪が細くなり、ハリ・コシがなくなってきた方"},"zh-TW":{"e":"頭髮一天比一天變細、缺乏彈力者"}}},{"k":"탈모 샴푸나 영양제만으로는 한계를 느끼시는 분","t":{"en":{"e":"Those who feel limits with only hair loss shampoo and nutrition"},"ja":{"e":"育毛シャンプーやサプリだけでは限界を感じている方"},"zh-TW":{"e":"僅靠防脫洗髮精或營養補充品感到效果有限者"}}},{"k":"스트레스, 다이어트, 산후 등으로 두피 환경이 무너지신 분","t":{"en":{"e":"Those whose scalp environment has collapsed due to stress, dieting, or postpartum changes"},"ja":{"e":"ストレス・ダイエット・産後などで頭皮環境が乱れている方"},"zh-TW":{"e":"因壓力、減肥、產後等導致頭皮環境受損者"}}},{"k":"단일 치료가 아닌, 체계적인 메디컬 관리가 필요하신 분","t":{"en":{"e":"Those who need systematic medical care rather than single treatments"},"ja":{"e":"単一治療ではなく、体系的な医療ケアが必要な方"},"zh-TW":{"e":"需要系統化醫療管理,而非單一療程者"}}},{"k":"3-Step 트리플 시너지, 왜 다를까요?","t":{"en":{"e":"3-Step Why is the triple synergy different"},"ja":{"e":"3-Step トリプルシナジー、なぜ違うのか?"},"zh-TW":{"e":"3-Step 三重協同作用,為何不同?"}}},{"k":"단순히 겉만 관리하는 것이 아닙니다. 탈모의 원인 차단부터 두피 속 섬유아세포를 자극하는 재생 단계까지, 의학적 근거를 바탕으로 모근의 힘을 길러줍니다.","t":{"en":{"e":"Not a simple management. It works precisely on each stage from the root cause of hair loss to the scalp environment. Based on medical evidence, it strengthens the power of hair roots"},"ja":{"e":"表面的なケアだけではありません。脱毛の原因を抑制する段階から、頭皮内部の線維芽細胞を刺激する再生段階まで、医学的根拠に基づき毛根の力を高めます。"},"zh-TW":{"e":"不只是表面護理。從阻斷落髮根本原因到刺激頭皮內纖維母細胞的再生階段,以醫學根據強化毛根力量"}}},{"k":"STEP 1 [차단] 메디컬 약물요법","t":{"en":{"e":"STEP 1 Medical medication therapy"},"ja":{"e":"STEP 1 【抑制】医療薬物療法"},"zh-TW":{"e":"STEP 1 【阻斷】醫療藥物療法"}}},{"k":"탈모를 유발하는 근본적인 원인을 의학적으로 억제하여, 더 이상의 진행을 막아주는 가장 탄탄한 기초 단계입니다.","t":{"en":{"e":"Medically suppresses the root causes that induce hair loss. Becomes a solid foundation stage that supports further progression"},"ja":{"e":"脱毛の根本原因を医学的に抑制し、進行を防ぐ最も重要な基礎段階です。"},"zh-TW":{"e":"以醫學方式抑制造成落髮的根本原因,阻止進一步惡化,是最穩固的基礎階段"}}},{"k":"STEP 2 [재생] 두피 메조테라피 주사","t":{"en":{"e":"STEP 2 Regeneration scalp mesotherapy injection"},"ja":{"e":"STEP 2 【再生】頭皮メソセラピー注射"},"zh-TW":{"e":"STEP 2 【再生】頭皮中胚層療法注射"}}},{"k":"모낭 생존에 필수적인 유효 성분(영양 물질)을 모근 가까이 직접 주사합니다. 두피 속부터 건강한 세포 재생을 유도하여 튼튼한 모발이 자랄 수 있는 토대를 만듭니다.","t":{"en":{"e":"Directly supplies essential active ingredients needed for hair growth. Rebuilds the scalp environment from the scalp. Helps hair to grow thicker"},"ja":{"e":"毛包の生存に必要な有効成分(栄養成分)を頭皮の深部に直接注入します。頭皮内部から促進健康な細胞再生から、強い毛髪が育つ土台を作ります。"},"zh-TW":{"e":"將毛囊存活所需的有效成分(營養物質)直接注入頭皮深層。從頭皮內部促進健康細胞再生,打造強韌毛髮生長基礎"}}},{"k":"STEP 3 [활성] LED 광에너지 테라피","t":{"en":{"e":"STEP 3 Activation LED care therapy"},"ja":{"e":"STEP 3 【活性】LED光エネルギー療法"},"zh-TW":{"e":"STEP 3 【活化】LED光能療法"}}},{"k":"특수 파장의 빛이 두피 깊숙이 침투하여 미세 혈액순환을 촉진합니다. 메조테라피 성분의 흡수율을 극대화하고, 두피 섬유아세포를 자극해 재생 시너지를 한층 더 높여줍니다.","t":{"en":{"e":"Special wavelength light penetrates deep into the scalp. Promotes regeneration of micro blood vessels and hair follicles. Maximizes mesotherapy effects and increases hair root cell regeneration activity"},"ja":{"e":"特殊な波長の光が頭皮の奥深くまで浸透し、微細な血行を促進します。メソセラピー成分の吸収を最大化し、線維芽細胞の活性化と再生シナジーをさらに高めます。"},"zh-TW":{"e":"特殊波長光深入頭皮,促進微循環,最大化中胚層療法成分吸收,提升頭皮纖維母細胞刺激與再生協同效果"}}},{"k":"시술 시간","t":{"en":{"e":"Procedure time"},"ja":{"e":"施術時間"},"zh-TW":{"e":"療程時間"}}},{"k":"약 30~40분 (빠르고 간편하게!)","t":{"en":{"e":"about 30 to 40 minutes / Fast and simple"},"ja":{"e":"約30~40分 (スピーディーで簡単)"},"zh-TW":{"e":"約30~40分鐘 (快速且便利)"}}},{"k":"마취 여부","t":{"en":{"e":"Anesthesia"},"ja":{"e":"麻酔の有無"},"zh-TW":{"e":"是否麻醉"}}},{"k":"불필요 (또는 연고 마취 진행)","t":{"en":{"e":"Not required / Topical anesthesia may be used"},"ja":{"e":"不要 (または麻酔クリーム使用)"},"zh-TW":{"e":"不需要 (或可進行表面麻醉)"}}},{"k":"일상 생활","t":{"en":{"e":"Daily life"},"ja":{"e":"日常生活"},"zh-TW":{"e":"日常生活"}}},{"k":"시술 직후 바로 일상생활 복귀 가능","t":{"en":{"e":"Return to daily activities possible immediately after the procedure"},"ja":{"e":"施術直後から日常生活に復帰可能"},"zh-TW":{"e":"療程後可立即恢復日常生活"}}},{"k":"추천 주기","t":{"en":{"e":"Recommended cycle"},"ja":{"e":"推奨周期"},"zh-TW":{"e":"建議週期"}}},{"k":"1~2주 간격 (개인별 두피 상태에 따라 맞춤 상담 진행)","t":{"en":{"e":"Every 1 to 2 weeks / Customized depending on individual scalp condition"},"ja":{"e":"1~2週間間隔 (頭皮状態に応じて個別カウンセリング)"},"zh-TW":{"e":"每1~2週一次 (依個人頭皮狀況進行客製化諮詢)"}}},{"k":"시크릿 이너볼륨 필러","t":{"en":{"e":"Secret Inner Volume Filler"},"ja":{"e":"シークレットインナーボリュームフィラー"},"zh":{"e":"秘密 内层 丰盈填充剂"},"zh-TW":{"e":"秘密 內層 豐盈填充劑"},"th":{"e":"ฟิลเลอร์ Secret Inner Volume"}},"c":"필러"},{"k":"5cc 715,000원","t":{"en":{"e":"5cc 715,000 KRW"},"ja":{"e":"5cc 715,000 KRW"},"zh":{"e":"5cc 715,000 KRW"},"zh-TW":{"e":"5cc 715,000 KRW"},"th":{"e":"5cc 715,000 KRW"}}},{"k":"속부터 차오르는 자신감","t":{"en":{"e":"Confidence that fills from within"},"ja":{"e":"内側から変わる、自信"},"zh":{"e":"从内在开始变化的自信感秘密"},"zh-TW":{"e":"從內在開始改變的自信感秘密"},"th":{"e":"ความมั่นใจที่เต็มเต็มจากภายใน"}}},{"k":"여성 인티메이트 전용 시술","t":{"en":{"e":"Intimate treatment for women"},"ja":{"e":"女性デリケートゾーン専用施術"},"zh":{"e":"女性私密部位专用施术"},"zh-TW":{"e":"女性私密部位專用療程"},"th":{"e":"หัตถการเฉพาะจุดซ่อนเร้นสำหรับผู้หญิง"}}},{"k":"만족도 UP!","t":{"en":{"e":"VOLUME UP!"},"ja":{"e":"ボリュームUP!"},"zh":{"e":"饱满UP!"},"zh-TW":{"e":"飽滿UP!"},"th":{"e":"ฉันเพิ่มได้!"}}},{"k":"자신감 UP!","t":{"en":{"e":"CONFIDENCE UP!"},"ja":{"e":"自信UP!"},"zh":{"e":"自信UP!"},"zh-TW":{"e":"自信UP!"},"th":{"e":"เพิ่มความมั่นใจ!"}}},{"k":"시크릿 이너볼륨 필러란?","t":{"en":{"e":"What is Secret Inner Volume Filler?"},"ja":{"e":"シークレットインナーボリュームフィラーとは?"},"zh":{"e":"什么是私密内在丰盈填充?"},"zh-TW":{"e":"秘密內在豐盈填充是什麼?"},"th":{"e":"ฟิลเลอร์เพิ่มวอลลุ่มภายในจุดซ่อนเร้นคืออะไร?"}}},{"k":"시크릿 이너볼륨 필러는 인체 친화적인 필러 성분을 사용하여 질 등의 여성 인티메이트 부위의 볼륨과 탄력을 자연스럽게 개선하는 시술입니다","t":{"en":{"e":"Secret Inner Volume Filler is a treatment that uses biocompatible filler ingredients to naturally improve volume and elasticity in women's intimate areas such as the vaginal region."},"ja":{"e":"シークレットインナーボリュームフィラーは、人体に優しいフィラー成分を使用し、膣などの女性のデリケートゾーンのボリュームと弾力を自然に改善する施術です"},"zh":{"e":"私密内在丰盈填充是一种采用对人体友好的填充成分,针对阴道等女性私密部位,自然改善其丰盈度与弹性的治疗方式。"},"zh-TW":{"e":"秘密內在豐盈填充 採用對人體友善的填充成分 針對陰道等女性私密部位 自然改善其豐盈度與彈性的療程"},"th":{"e":"ฟิลเลอร์เพิ่มวอลลุ่มภายในจุดซ่อนเร้น เป็นหัตถการที่ใช้ฟิลเลอร์ที่เป็นมิตรต่อร่างกาย เพื่อช่วยปรับวอลลุ่มและความยืดหยุ่นของจุดซ่อนเร้นของผู้หญิง เช่น บริเวณช่องคลอด ให้ดีขึ้นอย่างเป็นธรรมชาติ"}}},{"k":"빠른 효과로 눈에 보이고 몸으로 느껴지는 변화","t":{"en":{"e":"Fast results you can see with your eyes and feel with your body"},"ja":{"e":"即効性があり 目で見て分かり 身体で実感できる変化"},"zh":{"e":"快速见效 不仅看得见 身体也能感受到的变化"},"zh-TW":{"e":"快速見效 不只看得見 身體也能實際感受到的改變"},"th":{"e":"เห็นผลรวดเร็ว สังเกตได้ด้วยตา และสัมผัสได้จริง"}}},{"k":"확실한 탄력","t":{"en":{"e":"Firm elasticity"},"ja":{"e":"しっかりとした弾力"},"zh":{"e":"明显弹性"},"zh-TW":{"e":"明顯彈性"},"th":{"e":"ความกระชับชัดเจน"}}},{"k":"자연스러운 볼륨감","t":{"en":{"e":"Natural volume"},"ja":{"e":"自然なボリューム感"},"zh":{"e":"自然饱满感"},"zh-TW":{"e":"自然豐盈感"},"th":{"e":"วอลลุ่มดูเป็นธรรมชาติ"}}},{"k":"성감 극대화","t":{"en":{"e":"Enhanced confidence"},"ja":{"e":"自信の向上"},"zh":{"e":"自信提升"},"zh-TW":{"e":"自信提升"},"th":{"e":"เพิ่มความพึงพอใจในตัวเอง"}}},{"k":"미미썸 시크릿 볼륨 필러 추천 대상","t":{"en":{"e":"Secret Volume Filler Recommended for"},"ja":{"e":"ミミサム シークレットインナーボリュームフィラー おすすめの方"},"zh":{"e":"秘密内在丰盈填充 推荐对象"},"zh-TW":{"e":"秘密內在豐盈填充 推薦對象"},"th":{"e":"ผู้ที่เหมาะกับฟิลเลอร์ Secret Volume ของ Mimisome"}}},{"k":"성감을 보호하거나 개선하고 싶으신 분","t":{"en":{"e":"Those who want to enhance or improve sensitivity"},"ja":{"e":"出産や加齢により弾力の低下を感じる方"},"zh":{"e":"因分娩或老化导致弹性下降的人"},"zh-TW":{"e":"因生產或老化導致彈性下降者"},"th":{"e":"ผู้ที่ต้องการเสริมหรือปรับปรุงความรู้สึกขณะสัมผัส"}}},{"k":"파트너의 만족도를 빠르게 높이고 싶으신 분","t":{"en":{"e":"Those seeking quick volume satisfaction"},"ja":{"e":"デリケートゾーンのボリューム感を改善したい方"},"zh":{"e":"希望改善私密部位饱满感的人"},"zh-TW":{"e":"希望改善私密部位豐盈感者"},"th":{"e":"ผู้ที่ต้องการเพิ่ม Inner Volume อย่างรวดเร็ว"}}},{"k":"자연스럽게 자신감을 회복하고 싶으신 분","t":{"en":{"e":"Those who want to regain confidence naturally"},"ja":{"e":"自然に自信を取り戻したい方"},"zh":{"e":"希望自然恢复自信感的人"},"zh-TW":{"e":"希望自然恢復自信感者"},"th":{"e":"ผู้ที่ต้องการฟื้นคืนความมั่นใจให้เป็นธรรมชาติ"}}},{"k":"섬세한 맞춤 디자인","t":{"en":{"e":"Delicate Customized Design"},"ja":{"e":"繊細な触感の変化"},"zh":{"e":"细腻的触感改变"},"zh-TW":{"e":"細膩觸感的改變"},"th":{"e":"การออกแบบหัตถการอย่างละเอียด"}}},{"k":"안전성 + 지속력","t":{"en":{"e":"Safety + Longevity"},"ja":{"e":"安全性 + 持続性"},"zh":{"e":"安全性 + 持久性"},"zh-TW":{"e":"安全性 + 持久性"},"th":{"e":"ความปลอดภัย + ความคงทน"}}},{"k":"부드러운 히알루론산 성분으로 민감한 부위에도 안정적으로 적용되어 자연스러운 결과와 지속력을 제공합니다","t":{"en":{"e":"With soft hyaluronic acid ingredients, it can be safely applied even to sensitive areas, providing natural results and long-lasting effects."},"ja":{"e":"高い安定性を持つフィラーを使用し、日常生活に大きな影響をほとんど与えず、自然なボリュームと弾力の形成をサポートします。"},"zh":{"e":"使用高稳定性的填充材料,在不影响日常生活的情况下,帮助打造自然的容积与弹性。"},"zh-TW":{"e":"使用高穩定性的填充材料,在不影響日常生活的情況下,協助打造自然的容積與彈性。"},"th":{"e":"ด้วยสัมผัสที่นุ่มนวลและส่วนผสมที่เข้ากับร่างกาย จึงสามารถใช้กับบริเวณที่บอบบางได้อย่างปลอดภัย และให้ผลลัพธ์ที่ดูเป็นธรรมชาติ"}}},{"k":"약 20~30분 내외","t":{"en":{"e":"About 20-30 minutes"},"ja":{"e":"約20~30分"},"zh":{"e":"约20~30分钟"},"zh-TW":{"e":"約20~30分鐘"},"th":{"e":"ประมาณ 20-30 นาที"}}},{"k":"시술 후 빠르게 회복","t":{"en":{"e":"Quick recovery after treatment"},"ja":{"e":"施術後、比較的早く回復"},"zh":{"e":"施术后可快速恢复"},"zh-TW":{"e":"療程後可快速恢復"},"th":{"e":"กลับไปใช้ชีวิตประจำวันได้อย่างรวดเร็วหลังทำ"}}},{"k":"통증 정도","t":{"en":{"e":"Pain level"},"ja":{"e":"痛みの程度"},"zh":{"e":"疼痛程度"},"zh-TW":{"e":"疼痛程度"},"th":{"e":"ความเจ็บ"}}},{"k":"비교적 적은 편","t":{"en":{"e":"Relatively mild"},"ja":{"e":"比較的少なめ"},"zh":{"e":"相对较轻"},"zh-TW":{"e":"相對較低"},"th":{"e":"ค่อนข้างต่ำ"}}},{"k":"미미썸의원 시술 특징","t":{"en":{"e":"Key Features of MIMISOME Clinic Treatments"},"ja":{"e":"ミミソムクリニックの施術特徴"},"zh":{"e":"MIMISOME 诊所治疗特点"},"zh-TW":{"e":"MIMISOME 診所療程特色"},"th":{"e":"จุดเด่นของหัตถการ MIMISOME"}}},{"k":"1:1 맞춤상담","t":{"en":{"e":"1:1 personalized consultation"},"ja":{"e":"1:1 オーダーメイド施術"},"zh":{"e":"1:1 定制施术"},"zh-TW":{"e":"1:1 客製化療程"},"th":{"e":"ปรึกษาแบบตัวต่อตัว"}}},{"k":"노하우를 갖춘 의료진","t":{"en":{"e":"Experienced Medical Professionals"},"ja":{"e":"豊富な経験と高い技術力"},"zh":{"e":"丰富经验与技术力"},"zh-TW":{"e":"豐富經驗與技術實力"},"th":{"e":"ทีมแพทย์ผู้มีประสบการณ์"}}},{"k":"심미성 및 기능성 향상","t":{"en":{"e":"Enhanced Aesthetics and Functionality"},"ja":{"e":"機能性と美しさを同時に向上"},"zh":{"e":"功能与美感同步提升"},"zh-TW":{"e":"功能與美感同步提升"},"th":{"e":"เพิ่มทั้งความสวยงามและการใช้งาน"}}},{"k":"모공지우개 엑소좀 2.5cc","t":{"en":{"e":"Pore Care Exosome 2.5cc"},"ja":{"e":"毛穴ケア エクソソーム 2.5cc"},"zh-TW":{"e":"毛孔護理 外泌體 2.5cc"}},"c":"스킨부스터"},{"k":"2.5cc 99,000원","t":{"en":{"e":"2.5cc 99,000 KRW"},"ja":{"e":"2.5cc 99,000 KRW"},"zh-TW":{"e":"2.5cc 99,000 KRW"}}},{"k":"9.9만","t":{"en":{"e":"99,000 KRW"},"ja":{"e":"99,000 KRW"},"zh-TW":{"e":"99,000 KRW"}}},{"k":"MIMISOME EXOSOME","t":{"en":{"e":"MIMISOME EXOSOME"},"ja":{"e":"MIMISOME EXOSOME"},"zh-TW":{"e":"MIMISOME EXOSOME"}}},{"k":"근본적인 피부 개선 스킨부스터 엑소좀","t":{"en":{"e":"Fundamental Skin Improvement Skin Booster Exosome"},"ja":{"e":"根本からの肌改善 スキンブースター エクソソーム"},"zh-TW":{"e":"從根本改善膚況 外泌體肌膚導入療程"}}},{"k":"콜라겐 재생","t":{"en":{"e":"Collagen Care"},"ja":{"e":"コラーゲンケア"},"zh-TW":{"e":"膠原蛋白護理"}}},{"k":"유수분 밸런스","t":{"en":{"e":"Hydration Balance"},"ja":{"e":"水分バランス"},"zh-TW":{"e":"水分平衡"}}},{"k":"모공 축소","t":{"en":{"e":"Pore Care"},"ja":{"e":"毛穴ケア"},"zh-TW":{"e":"毛孔護理"}}},{"k":"미미썸 엑소좀 추천 대상","t":{"en":{"e":"MIMISOME Exosome Recommended For"},"ja":{"e":"ミミソム エクソソーム おすすめの方"},"zh-TW":{"e":"MIMISOME 外泌體 適合對象"}}},{"k":"거칠고 푸석한 피부결이 고민이신 분","t":{"en":{"e":"Those concerned about rough, dry skin"},"ja":{"e":"肌のざらつき・乾燥が気になる方"},"zh-TW":{"e":"在意肌膚粗糙、乾燥者"}}},{"k":"각종 트러블과 넓은 모공이 고민이신 분","t":{"en":{"e":"Those concerned about acne and enlarged pores"},"ja":{"e":"ニキビや毛穴の開きが気になる方"},"zh-TW":{"e":"在意痘痘與毛孔粗大者"}}},{"k":"복합적인 피부 개선을 원하시는 분","t":{"en":{"e":"Those who want to improve multiple skin concerns"},"ja":{"e":"複合的な肌悩みを改善したい方"},"zh-TW":{"e":"想改善多重肌膚問題者"}}},{"k":"무너진 피부 밸런스 케어","t":{"en":{"e":"Restore Your Skin Balance"},"ja":{"e":"崩れた肌バランスケア"},"zh-TW":{"e":"重整失衡的肌膚狀態"}}},{"k":"미미썸만의 입체적인 모공 토탈 케어","t":{"en":{"e":"MIMISOME's Signature Multi-Dimensional Pore Care"},"ja":{"e":"ミミソムだけの 立体的な毛穴トータルケア"},"zh-TW":{"e":"MIMISOME 專屬 立體毛孔全面護理"}}},{"k":"미미썸은 피부 상태를 정밀하게 분석하고 모공의 원인에 맞춰 시술을 설계합니다","t":{"en":{"e":"We carefully analyze your skin condition and design treatments based on the root causes of pores"},"ja":{"e":"ミミソムは肌状態を丁寧に分析し 毛穴の原因に合わせて施術を設計します"},"zh-TW":{"e":"透過細緻的膚況分析 依毛孔成因規劃客製療程"}}},{"k":"단순 축소가 아닌 피부결, 탄력, 밀도까지 함께 개선합니다","t":{"en":{"e":"Not just basic skincare, but a total approach to texture, elasticity, and glow"},"ja":{"e":"肌キメ・ハリ・ツヤまでトータルに整えます"},"zh-TW":{"e":"不只是基礎保養 從膚質、彈力到光澤全面調理"}}},{"k":"ASCE+ 핵심 성분 그대로","t":{"en":{"e":"ASCE+ Key Ingredients Intact"},"ja":{"e":"ASCE+ 主要成分そのまま"},"zh-TW":{"e":"ASCE+ 核心成分完整保留"}}},{"k":"본 이벤트는 침습적 방식으로 시술하지 않습니다","t":{"en":{"e":"This event does not involve invasive procedures"},"ja":{"e":"本イベントは侵襲的な方法では施術を行いません"},"zh-TW":{"e":"本活動非侵入式療程"}}},{"k":"ASCE+ 엑소좀 100%","t":{"en":{"e":"ASCE+ Exosome 100%"},"ja":{"e":"ASCE+ エクソソーム 100%"},"zh-TW":{"e":"ASCE+ 外泌體 100%"}}},{"k":"콜라겐 600%","t":{"en":{"e":"Collagen 600%"},"ja":{"e":"コラーゲン 600%"},"zh-TW":{"e":"膠原蛋白 600%"}}},{"k":"엘라스틴 300%","t":{"en":{"e":"Elastin 300%"},"ja":{"e":"エラスチン 300%"},"zh-TW":{"e":"彈力蛋白 300%"}}},{"k":"MIMISOME 오시는 길","t":{"en":{"e":"Directions to MIMISOME"},"ja":{"e":"MIMISOME へのアクセス"},"zh-TW":{"e":"MIMISOME 交通指南"}}},{"k":"진료시간안내","t":{"en":{"e":"Operating Hours"},"ja":{"e":"診療時間のご案内"},"zh-TW":{"e":"看診時間"}}},{"k":"평일 10:00 - 20:00","t":{"en":{"e":"Weekdays 10:00 - 20:00"},"ja":{"e":"平日 10:00 - 20:00"},"zh-TW":{"e":"平日 10:00 - 20:00"}}},{"k":"토요일 10:00 - 16:00","t":{"en":{"e":"Saturday 10:00 - 16:00"},"ja":{"e":"土曜日 10:00 - 16:00"},"zh-TW":{"e":"週六 10:00 - 16:00"}}},{"k":"점심시간 13:00 - 14:00","t":{"en":{"e":"Lunch break 13:00 - 14:00"},"ja":{"e":"ランチタイム 13:00 - 14:00"},"zh-TW":{"e":"午休時間 13:00 - 14:00"}}},{"k":"공휴일 / 일요일 휴진","t":{"en":{"e":"Closed on holidays and Sundays"},"ja":{"e":"祝日・日曜日休診"},"zh-TW":{"e":"國定假日／週日休診"}}},{"k":"팅커벨 귀필러","t":{"en":{"e":"Tinker Bell Ear Filler"},"ja":{"e":"ティンカーベル耳フィラー"},"zh":{"e":"小精灵耳部填充"},"zh-TW":{"e":"精靈耳玻尿酸"},"th":{"e":"ฉีดฟิลเลอร์หู"}},"c":"필러"},{"k":"팅커벨귀필러","t":{"en":{"e":"Tinker Bell Ear Filler"},"ja":{"e":"ティンカーベル耳フィラー"},"zh":{"e":"小精灵耳部填充"},"zh-TW":{"e":"精靈耳玻尿酸"},"th":{"e":"ฉีดฟิลเลอร์หู"}},"c":"필러"},{"k":"165,000원","t":{"en":{"e":"165,000 KRW"},"ja":{"e":"165,000 KRW"},"zh":{"e":"165,000 KRW"},"zh-TW":{"e":"165,000 KRW"},"th":{"e":"165,000 KRW"}}},{"k":"작아보이는 얼굴의 비결","t":{"en":{"e":"The secret to a smaller looking face"},"ja":{"e":"小顔に見せる秘訣"},"zh":{"e":"让脸看起来更小的秘诀"},"zh-TW":{"e":"讓臉看起來更小的秘密"},"th":{"e":"เคล็ดลับ ให้หน้าดูเล็กเวลาถ่ายรูป"}}},{"k":"비대칭 귓볼 & 얇은 귓볼","t":{"en":{"e":"Asymmetrical earlobes and thin earlobes"},"ja":{"e":"非対称の耳たぶと薄い耳たぶ"},"zh":{"e":"不对称的耳垂与薄耳垂"},"zh-TW":{"e":"不對稱耳垂與薄耳垂"},"th":{"e":"ติ่งหูไม่สมมาตร และติ่งหูบาง"}}},{"k":"칼 귀","t":{"en":{"e":"Sharp ears"},"ja":{"e":"尖った耳"},"zh":{"e":"尖耳"},"zh-TW":{"e":"尖耳"},"th":{"e":"หูแหลม"}}},{"k":"울퉁불퉁한 귀","t":{"en":{"e":"Uneven ears"},"ja":{"e":"不均一な耳"},"zh":{"e":"凹凸不平的耳朵"},"zh-TW":{"e":"不平整的耳朵"},"th":{"e":"หูไม่เรียบ"}}},{"k":"누운 귀","t":{"en":{"e":"Drooping ears"},"ja":{"e":"垂れた耳"},"zh":{"e":"贴头型耳朵"},"zh-TW":{"e":"貼頭型耳朵"},"th":{"e":"หูแนบศีรษะ"}}},{"k":"작은 귀","t":{"en":{"e":"Small ears"},"ja":{"e":"小さい耳"},"zh":{"e":"小耳朵"},"zh-TW":{"e":"小耳朵"},"th":{"e":"หูเล็ก"}}},{"k":"미미썸 팅커벨 귀필러를 통해 개선할 수 있습니다.","t":{"en":{"e":"Can be improved with MIMISOME Tinker Bell Ear Filler"},"ja":{"e":"ミミソムのティンカーベル耳フィラーで改善できます"},"zh":{"e":"可通过 Mimisome 小精灵耳部填充疗程改善"},"zh-TW":{"e":"精靈耳Mimisome 小精靈耳朵填充療程可改善"},"th":{"e":"สามารถปรับปรุงได้ด้วยฟิลเลอร์หูทิงเกอร์เบลล์ของ Mimisome"}}},{"k":"가장 이상적인 귀 비율","t":{"en":{"e":"Most ideal ear proportions"},"ja":{"e":"最も理想的な耳の比率"},"zh":{"e":"最理想的耳部比例"},"zh-TW":{"e":"最理想的耳朵比例"},"th":{"e":"สัดส่วนหูที่เหมาะสมที่สุด"}}},{"k":"예쁜 귀 길이","t":{"en":{"e":"Pretty length"},"ja":{"e":"美しい長さ"},"zh":{"e":"理想耳长"},"zh-TW":{"e":"理想耳長"},"th":{"e":"ความยาวหูที่สวยงาม"}}},{"k":"여성 5.5cm-6cm | 남성 7cm","t":{"en":{"e":"Women 5.5 cm to 6 cm, Men 7 cm"},"ja":{"e":"女性5.5cmから6cm, 男性7cm"},"zh":{"e":"女性 5.5-6 厘米 | 男性 7 厘米"},"zh-TW":{"e":"女性 5.5–6 公分 | 男性 7 公分"},"th":{"e":"ผู้หญิง 5.5-6 ซม. | ผู้ชาย 7 ซม."}}},{"k":"귓볼 크기","t":{"en":{"e":"Earlobe size"},"ja":{"e":"耳たぶの大きさ"},"zh":{"e":"耳垂大小"},"zh-TW":{"e":"耳垂大小"},"th":{"e":"ขนาดติ่งหู"}}},{"k":"귀 전체 크기의 25%","t":{"en":{"e":"25 percent of the entire ear"},"ja":{"e":"耳全体の25パーセント"},"zh":{"e":"约占整体耳朵的 25%"},"zh-TW":{"e":"約佔整體耳朵的 25%"},"th":{"e":"25% ของขนาดหูทั้งหมด"}}},{"k":"귀가 작고 누워있어 정면에서 보이지 않고 얼굴 비율이 마음에 들지 않는 경우","t":{"en":{"e":"If your ears are small and laid back, not visible from the front, and your facial proportions are not satisfying"},"ja":{"e":"耳が小さく後ろに寝ていて正面から見えにくい場合 顔のバランスに満足できない場合"},"zh":{"e":"若耳朵较小且贴头, 从正面不明显, 导致脸部比例不理想"},"zh-TW":{"e":"若耳朵較小且貼頭, 正面看不清楚, 臉部比例不理想時"},"th":{"e":"หากหูเล็ก แนบศีรษะ มองจากด้านหน้าไม่ค่อยเห็น และสัดส่วนใบหน้าไม่ถูกใจ"}}},{"k":"귀 모양 개선만으로 얼굴 길이와 크기 축소 효과 이목구비의 균형 및 얼굴을 조화롭게 만들어주는 시술입니다.","t":{"en":{"e":"This procedure improves ear shape to make the face look smaller and create more balanced facial proportions"},"ja":{"e":"耳の形を整え 顔をより小さく見せ バランスの取れたフェイスラインへ導きます"},"zh":{"e":"仅改善耳型 即可达到缩短脸部长度与比例的视觉效果 让五官更加平衡, 使整体脸型更协调"},"zh-TW":{"e":"僅改善耳型 即可達到縮短臉部長度與比例的視覺效果, 讓五官更加平衡, 使整體臉型更協調"},"th":{"e":"เพียงปรับรูปทรงหู สามารถช่วยให้ใบหน้าดูสั้นลงและเลดลง สร้างสมดุลของโครงหน้า และทำให้ใบหน้าดูสวยกลมกลืนมากขึ้น"}}},{"k":"소얼굴 효과 발군","t":{"en":{"e":"Outstanding face slimming effect"},"ja":{"e":"小顔効果抜群"},"zh":{"e":"小脸效果出众"},"zh-TW":{"e":"小顏效果拔群"},"th":{"e":"เห็นผลหน้าเล็กชัดเจน"}}},{"k":"화제의 팅커벨, 순식간에 예뻐지게 해드려요","t":{"en":{"e":"Trending Tinker Bell ears make you beautiful in an instant"},"ja":{"e":"話題のティンカーベル耳で一瞬で美しく"},"zh":{"e":"话题精灵耳 让你瞬间变美"},"zh-TW":{"e":"話題精靈耳 讓你瞬間變美"},"th":{"e":"ทิงเกอร์เบลล์สุดฮอต ทำให้คุณสวยในทันที"}}},{"k":"써마지 FLX 300샷","c":"리프팅","t":{"en":{"e":"Thermage FLX 300 Shots"},"ja":{"e":"サーマクールFLX 300ショット"},"zh-TW":{"e":"第四代鳳凰電波 300發"}}},{"k":"써마지 FLX","c":"리프팅","t":{"en":{"e":"Thermage FLX"},"ja":{"e":"サーマクールFLX"},"zh-TW":{"e":"第四代鳳凰電波"}}},{"k":"써마지FLX","c":"리프팅","t":{"en":{"e":"Thermage FLX"},"ja":{"e":"サーマクールFLX"},"zh-TW":{"e":"第四代鳳凰電波"}}},{"k":"피부 속부터 탱탱하게 차오르는","t":{"en":{"e":"Firming up from deep within the skin"},"ja":{"e":"肌の内側からふっくらと満ちる"},"zh-TW":{"e":"由肌底深層開始飽滿緊緻"}}},{"k":"고주파 리프팅","t":{"en":{"e":"RF Lifting"},"ja":{"e":"高周波リフティング"},"zh-TW":{"e":"電波拉提"}}},{"k":"고주파 에너지로 피부 탄력 개선","t":{"en":{"e":"Improves skin elasticity with radiofrequency energy"},"ja":{"e":"高周波エネルギーで肌のハリを改善"},"zh-TW":{"e":"透過電波能量改善肌膚彈性"}}},{"k":"피부 상태에 맞춘 맞춤 리프팅","t":{"en":{"e":"Customized lifting tailored to your skin condition"},"ja":{"e":"肌状態に合わせたオーダーメイドリフティング"},"zh-TW":{"e":"依照肌膚狀況打造客製化拉提"}}},{"k":"처진 윤곽과 잔주름까지 탄탄하게","t":{"en":{"e":"Helps firm sagging contours and smooth fine lines"},"ja":{"e":"たるんだフェイスラインや小ジワまでしっかりケア"},"zh-TW":{"e":"改善鬆弛輪廓與細紋,使肌膚更緊實"}}},{"k":"전방향 진동으로 통증 부담 완화","t":{"en":{"e":"Reduces discomfort with multidirectional vibration technology"},"ja":{"e":"全方向振動で痛みの負担を軽減"},"zh-TW":{"e":"以全方位震動技術降低疼痛負擔"}}},{"k":"무너진 탄력이 고민이라면?","t":{"en":{"e":"Concerned about loss of firmness?"},"ja":{"e":"失われたハリが気になるなら?"},"zh-TW":{"e":"如果您正煩惱肌膚彈力流失?"}}},{"k":"탄력개선","t":{"en":{"e":"Firming"},"ja":{"e":"ハリ改善"},"zh-TW":{"e":"彈力改善"}}},{"k":"윤곽정리","t":{"en":{"e":"Contour definition"},"ja":{"e":"輪郭補整"},"zh-TW":{"e":"輪廓緊緻"}}},{"k":"잔주름 완화","t":{"en":{"e":"Fine line reduction"},"ja":{"e":"小ジワ改善"},"zh-TW":{"e":"淡化細紋"}}},{"k":"스킨 타이트닝","t":{"en":{"e":"Skin tightening"},"ja":{"e":"スキンタイトニング"},"zh-TW":{"e":"肌膚緊實"}}},{"k":"이런 분들께 추천드려요","t":{"en":{"e":"Recommended for those who"},"ja":{"e":"このような方におすすめです"},"zh-TW":{"e":"推薦給這樣的您"}}},{"k":"얼굴 전체 탄력과 잔주름 개선을 원하시는 분","t":{"en":{"e":"Want to improve overall facial firmness and fine lines"},"ja":{"e":"顔全体のハリや小ジワを改善したい方"},"zh-TW":{"e":"想改善全臉彈力與細紋的人"}}},{"k":"리프팅 시술 시 통증이 부담스러우신 분","t":{"en":{"e":"Are concerned about discomfort during lifting treatments"},"ja":{"e":"リフティング施術の痛みが気になる方"},"zh-TW":{"e":"對拉提療程疼痛感到有負擔的人"}}},{"k":"늘어진 턱선과 얼굴라인이 고민이신 분","t":{"en":{"e":"Are concerned about a sagging jawline and facial contours"},"ja":{"e":"たるんだフェイスラインやあご周りが気になる方"},"zh-TW":{"e":"在意下顎線鬆弛與臉部輪廓下垂的人"}}},{"k":"입가·눈가 주름이 심해 나이 들어보이시는 분","t":{"en":{"e":"Have noticeable wrinkles around the mouth and eyes that make them look older"},"ja":{"e":"口元・目元のシワが目立ち、老けて見えやすい方"},"zh-TW":{"e":"因嘴角、眼周皺紋明顯而顯老的人"}}},{"k":"#탄력UP","t":{"en":{"e":"#Firmness UP"},"ja":{"e":"#ハリUP"},"zh-TW":{"e":"#彈力提升"}}},{"k":"#또렷한 윤곽","t":{"en":{"e":"#Defined contours"},"ja":{"e":"#くっきり輪郭"},"zh-TW":{"e":"#輪廓更明顯"}}},{"k":"#자연스러운 리프팅","t":{"en":{"e":"#Natural-looking lifting"},"ja":{"e":"#自然なリフティング"},"zh-TW":{"e":"#自然拉提效果"}}},{"k":"해당 사진은 미미썸의원의 광고용 이미지 모델입니다","t":{"en":{"e":"This image is for advertising purposes only for MIMISOME Clinic"},"ja":{"e":"本画像はミミサム医院の広告用イメージです"},"zh-TW":{"e":"本圖片僅作為米米森醫診所廣告形象示意使用"}}},{"k":"1,210,000원","t":{"en":{"e":"1,210,000 KRW"},"ja":{"e":"1,210,000 KRW"},"zh-TW":{"e":"1,210,000 KRW"}}},{"k":"VAT INCLUDED","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh-TW":{"e":"含VAT"}}},{"k":"*VAT INCLUDED","t":{"en":{"e":"*VAT INCLUDED"},"ja":{"e":"*VAT込み"},"zh-TW":{"e":"*含VAT"}}},{"k":"*VAT込み","t":{"en":{"e":"*VAT INCLUDED"},"ja":{"e":"*VAT込み"},"zh-TW":{"e":"*含VAT"}}},{"k":"*含VAT","t":{"en":{"e":"*VAT INCLUDED"},"ja":{"e":"*VAT込み"},"zh-TW":{"e":"*含VAT"}}},{"k":"포텐자 레이저","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza Laser"},"ja":{"e":"ポテンツァレーザー"}}},{"k":"포텐자레이저","c":"여드름/모공/흉터","t":{"en":{"e":"Potenza Laser"},"ja":{"e":"ポテンツァレーザー"}}},{"k":"미미썸 포텐자","t":{"en":{"e":"MIMISOME Potenza"},"ja":{"e":"ミミサム ポテンツァ"}}},{"k":"미미썸 포텐자 레이저","t":{"en":{"e":"MIMISOME Potenza Laser"},"ja":{"e":"ミミサム ポテンツァレーザー"}}},{"k":"#모공축소","t":{"en":{"e":"#PoreReduction"},"ja":{"e":"#毛穴縮小"}}},{"k":"#흉터치료","t":{"en":{"e":"#ScarTreatment"},"ja":{"e":"#傷跡治療"}}},{"k":"#탄력개선","t":{"en":{"e":"#ElasticityImprovement"},"ja":{"e":"#ハリ改善"}}},{"k":"얼굴에 여드름 흉터가 남았어","t":{"en":{"e":"I have acne scars left on my face"},"ja":{"e":"顔にニキビ跡が残っちゃった"}}},{"k":"나이가 들수록 모공이 안줄어드네","t":{"en":{"e":"My pores don't seem to shrink"},"ja":{"e":"年を重ねるごとに毛穴が目立つ"}}},{"k":"피부에 탄력 좀 주고 싶은데","t":{"en":{"e":"I want to improve my skin elasticity"},"ja":{"e":"肌にハリを出したいのに"}}},{"k":"어느새 생긴 모공&여드름흉터가 고민이라면","t":{"en":{"e":"If you're concerned about enlarged pores and acne scars that appeared over time"},"ja":{"e":"いつの間にかできた毛穴やニキビ跡が気になるなら"}}},{"k":"미미썸 포텐자로 해결해보세요!","t":{"en":{"e":"try solving it with Potenza!"},"ja":{"e":"ミミサム ポテンツァで解決してみてください!"}}},{"k":"미미썸 포텐자로 모공&여드름흉터 해결해요!","t":{"en":{"e":"Pores & acne scars? Potenza."},"ja":{"e":"ミミサム・ポテンツァで毛穴&ニキビ跡解決してみて!"}}},{"k":"미미썸 포텐자레이저는 고주파 종류와 깊이를 선택해 다양한 피부층에 에너지를 전달합니다","t":{"en":{"e":"Delivers energy to various skin layers by adjusting RF type and depth"},"ja":{"e":"ミミサム・ポテンツァレーザーは高周波の種類と深さを選択し、さまざまな皮膚層へエネルギーを届けます"}}},{"k":"시술횟수는 개개인의 피부 상태에 따라 상이하며, 적절한 시술횟수는 상담 후 확인하실 수 있습니다","t":{"en":{"e":"Sessions vary by skin condition. Final plan decided after consultation"},"ja":{"e":"施術回数は個人の肌状態によって異なり、適切な施術回数はカウンセリング後にご案内いたします"}}},{"k":"#깐달걀피부","t":{"en":{"e":"#GlassSkin"},"ja":{"e":"#ゆで卵肌"}}},{"k":"#자신감UP","t":{"en":{"e":"#ConfidenceUP"},"ja":{"e":"#自信アップ"}}},{"k":"#미미썸과상담","t":{"en":{"e":"#ConsultMIMISOME"},"ja":{"e":"#ミミサム相談"}}},{"k":"포텐자가 필요해","t":{"en":{"e":"You need Potenza"},"ja":{"e":"ポテンツァが必要かも"}}},{"k":"이런 분들에게 추천해드려요!","t":{"en":{"e":"Recommended for"},"ja":{"e":"こんな方におすすめです!"}}},{"k":"피부 탄력이 떨어져 잔주름이 고민이신 분","t":{"en":{"e":"Concerned about wrinkles from loss of elasticity"},"ja":{"e":"肌のハリが低下し・小ジワが気になる方"}}},{"k":"자연스러운 리프팅, 타이트닝을 원하시는 분","t":{"en":{"e":"Those who want natural lifting and tightening"},"ja":{"e":"自然なリフトアップ・タイトニングを求める方"}}},{"k":"얼굴 라인, 모공, 피부결이 고민이신 분","t":{"en":{"e":"Concerned about contour, pores, and texture"},"ja":{"e":"フェイスライン・毛穴・キメが気になる方"}}},{"k":"예민해진 피부를 재생시키고 싶으신 분","t":{"en":{"e":"Those who want to restore sensitive or damaged skin"},"ja":{"e":"敏感になった肌を再生したい方"}}},{"k":"*카운셀링 포함","t":{"en":{"e":"*CONSULTATION INCLUDED"},"ja":{"e":"※カウンセリング込み"}}},{"k":"※カウンセリング込み","t":{"en":{"e":"*CONSULTATION INCLUDED"},"ja":{"e":"※カウンセリング込み"}}},{"k":"※VAT込み","t":{"en":{"e":"*VAT INCLUDED"},"ja":{"e":"※VAT込み"}}},{"k":"209,000원","t":{"en":{"e":"209,000 KRW"},"ja":{"e":"209,000 KRW"}}},{"k":"리쥬란 HB PLUS","t":{"en":{"e":"REJURAN HB PLUS"},"ja":{"e":"リジュラン HB PLUS"},"zh":{"e":"丽珠兰 HB PLUS"},"zh-TW":{"e":"麗珠蘭 HB PLUS"},"th":{"e":"REJURAN HB PLUS"}},"c":"스킨부스터"},{"k":"리쥬란 HB plus","t":{"en":{"e":"REJURAN HB plus"},"ja":{"e":"リジュラン HB plus"},"zh":{"e":"丽珠兰 HB plus"},"zh-TW":{"e":"麗珠蘭 HB plus"},"th":{"e":"รีจูรัน HB plus"}},"c":"스킨부스터"},{"k":"리쥬란HB plus","t":{"en":{"e":"REJURAN HB plus"},"ja":{"e":"リジュランHB plus"},"zh":{"e":"丽珠兰 HB plus"},"zh-TW":{"e":"麗珠蘭 HB plus"},"th":{"e":"รีจูรัน HB plus"}},"c":"스킨부스터"},{"k":"리쥬란HB+","t":{"en":{"e":"Rejuran HB+"},"ja":{"e":"リジュランHB+"},"zh":{"e":"丽珠兰 HB+"},"zh-TW":{"e":"麗珠蘭 HB+"},"th":{"e":"รีจูรัน HB+"}},"c":"스킨부스터"},{"k":"피부 속부터 탄탄하게","t":{"en":{"e":"Firming from deep within the skin"},"ja":{"e":"肌の内側からしっかりと整える"},"zh":{"e":"从肌肤深层开始紧致"},"zh-TW":{"e":"從肌膚深層開始緊緻"},"th":{"e":"ฟื้นฟูผิวให้แข็งแรงจากชั้นใน"}}},{"k":"가장 근본적인 안티에이징!","t":{"en":{"e":"The most fundamental anti-aging solution"},"ja":{"e":"根本的なアンチエイジング"},"zh":{"e":"根本性抗衰老"},"zh-TW":{"e":"根本性抗老化"},"th":{"e":"การชะลอวัยจากต้นตอ"}}},{"k":"피부 깊은 곳에서부터 차오르는 스킨부스터","t":{"en":{"e":"A skin booster that revitalizes the skin from deep within"},"ja":{"e":"肌の奥深くから満ちあふれるスキンブースター"},"zh":{"e":"从肌肤深层焕发的肌肤增强针"},"zh-TW":{"e":"從肌膚深層湧現的肌膚強化針"},"th":{"e":"สกินบูสเตอร์ที่ฟื้นฟูจากชั้นผิวลึก"}}},{"k":"#피부재생","t":{"en":{"e":"#Skin regeneration"},"ja":{"e":"#皮膚再生"},"zh":{"e":"#皮肤再生"},"zh-TW":{"e":"#皮膚再生"},"th":{"e":"#ฟื้นฟูผิว"}}},{"k":"#탄력상승","t":{"en":{"e":"#Improved elasticity"},"ja":{"e":"#弾力アップ"},"zh":{"e":"#弹力提升"},"zh-TW":{"e":"#彈力提升"},"th":{"e":"#เพิ่มความกระชับ"}}},{"k":"#피부결 개선","t":{"en":{"e":"#Improved skin texture"},"ja":{"e":"#肌キメ改善"},"zh":{"e":"#肤质改善"},"zh-TW":{"e":"#膚質改善"},"th":{"e":"#ผิวเรียบเนียนขึ้น"}}},{"k":"얇아진 피부, 피부 속부터 개선","t":{"en":{"e":"Improves thinned skin from within"},"ja":{"e":"薄くなった肌 肌の内側から改善"},"zh":{"e":"变薄的肌肤，从肌肤内部开始改善"},"zh-TW":{"e":"變薄的肌膚，從肌膚內部開始改善"},"th":{"e":"ผิวที่บางลง ฟื้นฟูจากภายในผิว"}}},{"k":"약해진 피부 속 깊은 곳부터 차오르는 재생","t":{"en":{"e":"Promotes regeneration deep within weakened skin"},"ja":{"e":"弱くなった肌の奥深くから 湧き上がる再生"},"zh":{"e":"脆弱的肌肤内部，从深层涌现的新生"},"zh-TW":{"e":"脆弱的肌膚內部，從深層湧現的新生"},"th":{"e":"ผิวที่อ่อนแอ ฟื้นฟูจากชั้นลึกภายใน"}}},{"k":"깨끗하고 맑은 피부로 돌아가요!","t":{"en":{"e":"Helps restore clear, healthy-looking skin!"},"ja":{"e":"クリアで透明感のある肌へ戻りましょう！"},"zh":{"e":"重回洁净透亮的肌肤！"},"zh-TW":{"e":"重回潔淨透亮的肌膚！"},"th":{"e":"กลับสู่ผิวใสสะอาดอีกครั้ง!"}}},{"k":"리쥬란이란?","t":{"en":{"e":"What is REJURAN?"},"ja":{"e":"リジュランHB+とは？"},"zh":{"e":"什么是丽珠兰？"},"zh-TW":{"e":"什麼是麗珠蘭？"},"th":{"e":"รีจูรันคืออะไร?"}}},{"k":"연어주사라고 불리는 리쥬란은 연어에서 추출한 폴리뉴클레오타이드(PN)성분을 피부에 주입해 피부 속 환경을 개선해 주는 스킨부스터 시술입니다.","t":{"en":{"e":"Rejuran, also known as the \\"salmon injection,\\" is a skin booster treatment that uses polynucleotide (PN) extracted from salmon to improve the skin's internal environment."},"ja":{"e":"サーモン由来のポリヌクレオチド(PN)成分を肌に注入し、肌の内側の環境を改善するスキンブースターです。"},"zh":{"e":"将从鲑鱼中提取的多核苷酸 (PN) 成分注入肌肤，改善肌肤内部环境的肌肤增强针。"},"zh-TW":{"e":"將從鮭魚中萃取的多核苷酸 (PN) 成分注入肌膚，改善肌膚內部環境的肌膚強化針。"},"th":{"e":"สกินบูสเตอร์ที่ฉีดสารโพลีนิวคลีโอไทด์ (PN) ซึ่งสกัดจากปลาแซลมอนเข้าสู่ผิว เพื่อฟื้นฟูสภาพผิวจากภายใน"}}},{"k":"인체에 안전한 생체 적합 물질 PN 성분","t":{"en":{"e":"Biocompatible PN ingredients that are safe for the human body"},"ja":{"e":"人体に安全な生体適合PN成分"},"zh":{"e":"对人体安全的生物相容PN成分"},"zh-TW":{"e":"對人體安全的生物相容PN成分"},"th":{"e":"สาร PN ที่เข้ากันได้กับร่างกายและปลอดภัย"}}},{"k":"일정한 형태를 유지하며 서서히 분해","t":{"en":{"e":"Maintains a consistent structure and gradually breaks down over time"},"ja":{"e":"一定の形状を保ちながらゆっくり分解"},"zh":{"e":"保持固定形态并逐渐分解"},"zh-TW":{"e":"保持固定形態並逐漸分解"},"th":{"e":"คงรูปทรงและสลายตัวอย่างช้าๆ"}}},{"k":"통증은 Down","t":{"en":{"e":"LESS PAIN Down"},"ja":{"e":"痛みは Down"},"zh":{"e":"疼痛是 Down"},"zh-TW":{"e":"疼痛是 Down"},"th":{"e":"ความเจ็บปวดคือ Down"}}},{"k":"효과는 UP","t":{"en":{"e":"ENHANCED RESULTS UP"},"ja":{"e":"効果は UP"},"zh":{"e":"效果是 UP"},"zh-TW":{"e":"效果是 UP"},"th":{"e":"ผลลัพธ์คือ UP"}}},{"k":"기존 리쥬란 힐러의 히알루론산과 국소마취성분(리도카인)이 더해져 통증은 줄이고 수분감은 업그레이드 시켰습니다.","t":{"en":{"e":"With the addition of hyaluronic acid and the local anesthetic lidocaine to the original Rejuran Healer formula, discomfort is reduced while hydration is enhanced."},"ja":{"e":"リジュランシリーズにヒアルロン酸と局所麻酔成分リドカインを配合し、痛みを軽減し保湿力をアップグレードしました。"},"zh":{"e":"在原有丽珠兰系列基础上添加透明质酸，并加入局部麻醉成分利多卡因，降低疼痛感并升级保湿效果。"},"zh-TW":{"e":"在原有麗珠蘭系列基礎上添加玻尿酸，並加入局部麻醉成分利多卡因，降低疼痛感並升級保濕效果。"},"th":{"e":"เพิ่มกรดไฮยาลูรอนิกและสารยาชาเฉพาะที่ลิโดเคนในรีจูรันเดิม ช่วยลดความเจ็บและเพิ่มความชุ่มชื้นให้ดียิ่งขึ้น"}}},{"k":"히알루론산","t":{"en":{"e":"Hyaluronic Acid"},"ja":{"e":"ヒアルロン酸"},"zh":{"e":"透明质酸"},"zh-TW":{"e":"玻尿酸"},"th":{"e":"กรดไฮยาลูรอนิก"}}},{"k":"리쥬란 힐러","t":{"en":{"e":"Rejuran Healer"},"ja":{"e":"リジュラン成分"},"zh":{"e":"丽珠兰成分"},"zh-TW":{"e":"麗珠蘭成分"},"th":{"e":"สารรีจูรัน"}}},{"k":"리도카인","t":{"en":{"e":"Lidocaine"},"ja":{"e":"リドカイン"},"zh":{"e":"利多卡因"},"zh-TW":{"e":"利多卡因"},"th":{"e":"ลิโดเคน"}}},{"k":"이런 분들께 추천합니다!","t":{"en":{"e":"Recommended for those who!"},"ja":{"e":"こんな方におすすめです！"},"zh":{"e":"推荐给以下人群！"},"zh-TW":{"e":"推薦給以下族群！"},"th":{"e":"เหมาะสำหรับผู้ที่มีปัญหาเหล่านี้!"}}},{"k":"피부가 건조하고 푸석해진 분","t":{"en":{"e":"Have dry, rough skin"},"ja":{"e":"肌が乾燥してカサついている方"},"zh":{"e":"皮肤干燥粗糙的人"},"zh-TW":{"e":"肌膚乾燥粗糙的人"},"th":{"e":"ผู้ที่ผิวแห้งและหยาบกร้าน"}}},{"k":"탄력 저하와 잔주름이 생긴 분","t":{"en":{"e":"Are experiencing loss of elasticity and fine lines"},"ja":{"e":"弾力低下や小じわが気になる方"},"zh":{"e":"出现弹力下降和细纹的人"},"zh-TW":{"e":"出現彈力下降與細紋的人"},"th":{"e":"ผู้ที่มีความกระชับลดลงและริ้วรอยเล็กๆ"}}},{"k":"윤기와 생기가 부족한 피부인 분","t":{"en":{"e":"Have dull, tired-looking skin lacking radiance and vitality"},"ja":{"e":"ツヤとハリが不足している肌の方"},"zh":{"e":"缺乏光泽与活力的肌肤"},"zh-TW":{"e":"缺乏光澤與活力的肌膚"},"th":{"e":"ผู้ที่ผิวขาดความเปล่งปลั่งและความสดใส"}}},{"k":"모공이 넓고 피부결이 거친 분","t":{"en":{"e":"Have enlarged pores and uneven skin texture"},"ja":{"e":"毛穴が目立ち、肌キメが荒れている方"},"zh":{"e":"毛孔粗大、肤质粗糙的人"},"zh-TW":{"e":"毛孔粗大、膚質粗糙的人"},"th":{"e":"ผู้ที่มีรูขุมขนกว้างและผิวไม่เรียบ"}}},{"k":"보습과 탄력을 동시에 원하는 분","t":{"en":{"e":"Want both hydration and improved elasticity"},"ja":{"e":"保湿と弾力を同時に求める方"},"zh":{"e":"希望同时改善保湿与弹力的人"},"zh-TW":{"e":"希望同時改善保濕與彈力的人"},"th":{"e":"ผู้ที่ต้องการเพิ่มทั้งความชุ่มชื้นและความกระชับพร้อมกัน"}}},{"k":"99,000원","t":{"en":{"e":"99,000 KRW"},"ja":{"e":"99,000 KRW"},"zh":{"e":"99,000 KRW"},"zh-TW":{"e":"99,000 KRW"},"th":{"e":"99,000 KRW"}}},{"k":"含增值税","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"含增值稅","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"รวม VAT แล้ว","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含增值稅"},"th":{"e":"รวม VAT แล้ว"}}},{"k":"쥬베룩","t":{"en":{"e":"JUVELOOK / Juvelook"},"ja":{"e":"ジュベルック"},"zh":{"e":"乔雅露"},"zh-TW":{"e":"喬雅露"},"th":{"e":"Juvelook"}},"c":"스킨부스터"},{"k":"쥬베룩 3cc","t":{"en":{"e":"JUVELOOK 3cc"},"ja":{"e":"ジュベルック 3cc"},"zh":{"e":"乔雅露 3cc"},"zh-TW":{"e":"喬雅露 3cc"},"th":{"e":"Juvelook 3cc"}},"c":"스킨부스터"},{"k":"쥬베룩 스킨 3cc","t":{"en":{"e":"JUVELOOK Skin 3cc"},"ja":{"e":"ジュベルックスキン 3cc"},"zh":{"e":"乔雅露 3cc"},"zh-TW":{"e":"JUVELOOK喬雅露 3cc"},"th":{"e":"JUVELOOK 3cc"}},"c":"스킨부스터"},{"k":"피부를 건강하고 탄력있게 만들어주는","t":{"en":{"e":"Promotes healthier, firmer-looking skin"},"ja":{"e":"肌を健康でハリのある状態へ導く"},"zh":{"e":"打造健康紧致肌肤的"},"zh-TW":{"e":"讓肌膚健康有彈力的"},"th":{"e":"ช่วยให้ผิวดูสุขภาพดีและกระชับ"}}},{"k":"자가콜라겐 생성 촉진","t":{"en":{"e":"Stimulates natural collagen production"},"ja":{"e":"自己コラーゲン生成促進"},"zh":{"e":"促进自身胶原蛋白生成"},"zh-TW":{"e":"促進自身膠原蛋白生成"},"th":{"e":"กระตุ้นการสร้างคอลลาเจนตามธรรมชาติ"}}},{"k":"피부 탄력 및 잔주름 개선, 모공 개선","t":{"en":{"e":"Helps improve skin elasticity, fine lines, and the appearance of pores"},"ja":{"e":"肌のハリ向上・小じわ改善・毛穴改善"},"zh":{"e":"改善皮肤弹力与细纹，改善毛孔"},"zh-TW":{"e":"改善肌膚彈力與細紋，改善毛孔"},"th":{"e":"ความกระชับของผิวและลดริ้วรอยเล็กๆ"}}},{"k":"탄력과 물광을 동시에 원하시나요?","t":{"en":{"e":"Are you looking for both firmness and a radiant glow?"},"ja":{"e":"ハリとツヤを同時に求めていますか？"},"zh":{"e":"想同时拥有紧致与水光肌吗？"},"zh-TW":{"e":"想同時擁有彈力與水光肌嗎？"},"th":{"e":"ต้องการทั้งความกระชับและผิวฉ่ำโกลว์พร้อมกันหรือไม่?"}}},{"k":"피부재생","t":{"en":{"e":"Skin rejuvenation"},"ja":{"e":"肌再生"},"zh":{"e":"肌肤再生"},"zh-TW":{"e":"肌膚再生"},"th":{"e":"ฟื้นฟูผิว"}}},{"k":"잔주름 개선","t":{"en":{"e":"Improvement of fine lines"},"ja":{"e":"小じわ改善"},"zh":{"e":"改善细纹"},"zh-TW":{"e":"改善細紋"},"th":{"e":"ลดริ้วรอยเล็กๆ"}}},{"k":"자가콜라겐 생성","t":{"en":{"e":"Natural collagen stimulation"},"ja":{"e":"自己コラーゲン生成"},"zh":{"e":"促进胶原生成"},"zh-TW":{"e":"促進膠原生成"},"th":{"e":"คอลลาเจนตามธรรมชาติถูกกระตุ้น"}}},{"k":"넓은 모공 개선","t":{"en":{"e":"Improvement of enlarged pores"},"ja":{"e":"開いた毛穴改善"},"zh":{"e":"改善粗大毛孔"},"zh-TW":{"e":"改善粗大毛孔"},"th":{"e":"ช่วยกระชับรูขุมขนกว้าง"}}},{"k":"이런 분들에게 추천해요.","t":{"en":{"e":"Recommended for those who"},"ja":{"e":"こんな方におすすめです."},"zh":{"e":"推荐给以下人群."},"zh-TW":{"e":"推薦給這些人."},"th":{"e":"แนะนำสำหรับผู้ที่"}}},{"k":"건강한 피부를 원하시는 분","t":{"en":{"e":"Want healthier-looking skin"},"ja":{"e":"健康的な肌を目指したい方"},"zh":{"e":"希望拥有健康肌肤的人"},"zh-TW":{"e":"想要健康肌膚的人"},"th":{"e":"ผู้ที่ต้องการผิวสุขภาพดี"}}},{"k":"피부가 푸석하고 거칠며, 탄력이 떨어지신 분","t":{"en":{"e":"Have dull, rough skin with reduced elasticity"},"ja":{"e":"肌が乾燥してごわつき、ハリが低下している方"},"zh":{"e":"皮肤干燥粗糙、弹力下降的人"},"zh-TW":{"e":"肌膚乾燥粗糙、彈力下降的人"},"th":{"e":"ผู้ที่มีผิวแห้งหยาบและความกระชับลดลง"}}},{"k":"자글한 잔주름이 고민이신 분","t":{"en":{"e":"Are concerned about fine lines"},"ja":{"e":"細かいしわが気になる方"},"zh":{"e":"为细小皱纹烦恼的人"},"zh-TW":{"e":"為細小皺紋困擾的人"},"th":{"e":"ผู้ที่กังวลเรื่องริ้วรอยเล็กๆ"}}},{"k":"보습/탄력/재생 등 복합 시술을 원하시는 분","t":{"en":{"e":"Are looking for a treatment that addresses hydration, firmness, and skin renewal all at once"},"ja":{"e":"保湿・ハリ・再生など総合的なケアを希望する方"},"zh":{"e":"希望进行保湿/紧致/修复等综合护理的人"},"zh-TW":{"e":"希望進行保濕／彈力／修復等複合療程的人"},"th":{"e":"ผู้ที่ต้องการการดูแลแบบครบทั้งความชุ่มชื้น ความกระชับ และการฟื้นฟู"}}},{"k":"*생성형 AI 모델입니다.","t":{"en":{"e":"*THIS IS A GENERATIVE AI MODEL"},"ja":{"e":"*生成型AIモデルです"},"zh":{"e":"*生成式AI模型"},"zh-TW":{"e":"*生成式AI模型"},"th":{"e":"*เป็นภาพโมเดลที่สร้างจาก AI"}}},{"k":"생성형 AI 모델입니다.","t":{"en":{"e":"THIS IS A GENERATIVE AI MODEL"},"ja":{"e":"生成型AIモデルです"},"zh":{"e":"生成式AI模型"},"zh-TW":{"e":"生成式AI模型"},"th":{"e":"เป็นภาพโมเดลที่สร้างจาก AI"}}},{"k":"275,000원","t":{"en":{"e":"275,000 KRW"},"ja":{"e":"275,000 KRW"},"zh":{"e":"27.5万韩元"},"zh-TW":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"27.5만원","t":{"en":{"e":"275,000 KRW"},"ja":{"e":"275,000 KRW"},"zh":{"e":"27.5万韩元"},"zh-TW":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"27.5만","t":{"en":{"e":"275,000 KRW"},"ja":{"e":"275,000 KRW"},"zh":{"e":"27.5万韩元"},"zh-TW":{"e":"27.5萬韓元"},"th":{"e":"275,000 KRW"}}},{"k":"*상담 후 결정","t":{"en":{"e":"*Subject to consultation"},"ja":{"e":"※カウンセリング込み"},"zh":{"e":"需先咨询"},"zh-TW":{"e":"需先諮詢"},"th":{"e":"ปรึกษาก่อน"}}},{"k":"※需先咨询","t":{"en":{"e":"*Consultation required first"},"ja":{"e":"※カウンセリング後にご案内"},"zh":{"e":"※需先咨询"},"zh-TW":{"e":"※需先諮詢"},"th":{"e":"※ต้องปรึกษาก่อน"}}},{"k":"※需先諮詢","t":{"en":{"e":"*Consultation required first"},"ja":{"e":"※カウンセリング後にご案内"},"zh":{"e":"※需先咨询"},"zh-TW":{"e":"※需先諮詢"},"th":{"e":"※ต้องปรึกษาก่อน"}}},{"k":"含税","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"含稅","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"รวมภาษีมูลค่าเพิ่มแล้ว","t":{"en":{"e":"VAT INCLUDED"},"ja":{"e":"VAT込み"},"zh":{"e":"含增值税"},"zh-TW":{"e":"含增值稅"},"th":{"e":"รวมภาษีมูลค่าเพิ่มแล้ว"}}},{"k":"무턱필러 + 슈링크 100샷","t":{"en":{"e":"Chin Filler + Shrink 100 Shots"},"zh-TW":{"e":"下巴填充 + Shurink 100發"}},"c":"필러"},{"k":"무턱필러+슈링크 100샷","t":{"en":{"e":"Chin Filler + SHRINK 100 Shots"},"zh-TW":{"e":"下巴填充 + SHRINK 100發"}},"c":"필러"},{"k":"무턱필러","t":{"en":{"e":"Chin Filler"},"zh-TW":{"e":"下巴填充"}},"c":"필러"},{"k":"#인상개선","t":{"en":{"e":"#ImpressionImprovement"},"zh-TW":{"e":"#小臉線條"}}},{"k":"#V라인얼굴","t":{"en":{"e":"#VLineFace"},"zh-TW":{"e":"#V型下巴"}}},{"k":"#얼굴전체탄력","t":{"en":{"e":"#FullFaceElasticity"},"zh-TW":{"e":"#臉部比例平衡"}}},{"k":"무턱","t":{"en":{"e":"recessed chin"},"zh-TW":{"e":"無下巴"}}},{"k":"돌출입","t":{"en":{"e":"protruding mouth"},"zh-TW":{"e":"下巴前凸"}}},{"k":"뭉툭한턱","t":{"en":{"e":"blunt chin"},"zh-TW":{"e":"厚重下巴"}}},{"k":"하관이 고민이라면?","t":{"en":{"e":"Concerned about your lower face?"},"zh-TW":{"e":"如果你在意下半臉？"}}},{"k":"짧고 볼륨이 부족한 턱 끝에 얼굴 전체적인 밸런스가 맞도록 디자인하여 보다 더 갸름하고 세련된 인상으로 개선","t":{"en":{"e":"Designs the short, volume-deficient chin tip to balance the overall face, improving impression to a slimmer and more refined look"},"zh-TW":{"e":"針對下巴較短或下巴末端飽滿度不足的情況 從整體臉部比例出發 進行精細設計 改善為更加立體、精緻的臉部輪廓"}}},{"k":"턱 아래 늘어진 피부를 탄탄하게 잡아주고 흐릿해진 턱선을 위로 끌어올려 윤곽을 또렷하게 정리하며 하관 라인을 매끄럽고 선명하게 완성","t":{"en":{"e":"Firmly tightens sagging skin under the chin and lifts a blurred jawline upward to define the contour, completing a smooth and clear lower face line"},"zh-TW":{"e":"緊實下巴下方 鬆弛的肌膚 將鬆散的輪廓線 向上拉提 讓線條更加俐落 打造俐落分明的下半臉輪廓"}}},{"k":"무턱 때문에 얼굴이 둥글거나 짧아 보이는 경우","t":{"en":{"e":"Cases where the face looks round or short due to a recessed chin"},"zh-TW":{"e":"因下巴問題 臉型看起來較圓或較短的情況"}}},{"k":"턱선이 흐려서 이중턱이 강조되는 경우","t":{"en":{"e":"Cases where a blurred jawline emphasizes a double chin"},"zh-TW":{"e":"下顎線鬆弛 雙下巴明顯的情況"}}},{"k":"필러만 하면 턱이 부자연스러울까 걱정되는 분","t":{"en":{"e":"Those worried that filler alone might look unnatural"},"zh-TW":{"e":"僅進行下巴填充 擔心臉部比例失衡的人"}}},{"k":"수술 없이 하관 라인 정리를 원하시는 분","t":{"en":{"e":"Those who want to refine the lower face line without surgery"},"zh-TW":{"e":"希望不動手術 整理下顎線條的人"}}},{"k":"대용량 갸름주사","t":{"en":{"e":"Large-Capacity Slimming Injection"},"zh-TW":{"e":"大容量 臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"얼굴지방분해주사","t":{"en":{"e":"Face Fat Dissolve Injection"},"zh-TW":{"e":"臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"갸름주사","t":{"en":{"e":"Slimming Injection"},"zh-TW":{"e":"臉部溶脂注射"}},"c":"비만/다이어트"},{"k":"#숨어있던","t":{"en":{"e":"#Hidden"},"zh-TW":{"e":"#隱藏的"}}},{"k":"#V라인찾기","t":{"en":{"e":"#FindVLine"},"zh-TW":{"e":"#尋找V型臉"}}},{"k":"FACE FAT DISSOLVE","t":{"en":{"e":"FACE FAT DISSOLVE"},"zh-TW":{"e":"FACE FAT DISSOLVE"}}},{"k":"단순한 주사 그 이상, 얼굴라인을 되찾다","t":{"en":{"e":"More than just an injection — reclaim your face line"},"zh-TW":{"e":"不只是打一針，找回清晰的臉部線條"}}},{"k":"이중턱 및 볼살 감소","t":{"en":{"e":"Reduces double chin and cheek fat"},"zh-TW":{"e":"減少雙下巴與臉頰脂肪"}}},{"k":"1:1 맞춤 시술","t":{"en":{"e":"1:1 customized treatment"},"zh-TW":{"e":"1:1 客製化療程"}}},{"k":"조화로운 얼굴","t":{"en":{"e":"Harmonious facial proportions"},"zh-TW":{"e":"協調的臉部比例"}}},{"k":"V라인 얼굴형","t":{"en":{"e":"V-line face shape"},"zh-TW":{"e":"V型臉型"}}},{"k":"단순히 빼는 시대는 끝!","t":{"en":{"e":"The era of simply removing fat is over!"},"zh-TW":{"e":"只追求減少的時代已結束！"}}},{"k":"섬세하고 디테일한 몰딩","t":{"en":{"e":"Delicate and detailed molding"},"zh-TW":{"e":"細緻且講究細節的塑形"}}},{"k":"얼굴형을 고려한","t":{"en":{"e":"Tailored to your face shape"},"zh-TW":{"e":"依臉型設計的"}}},{"k":"1:1 맞춤 디자인","t":{"en":{"e":"1:1 customized design"},"zh-TW":{"e":"1:1 客製化設計"}}},{"k":"미미썸 얼굴지방분해주사는 얼굴형과 지방 분포를 고려한 1:1 디자인 시술로 보다 입체적이고 조화로운 작은 얼굴을 연출합니다.","t":{"en":{"e":"MIMISOME face fat dissolve injection considers face shape and fat distribution with 1:1 designed treatment to create a more three-dimensional, harmonious small face."},"zh-TW":{"e":"MIMISOME 臉部溶脂注射考量臉型與脂肪分布 進行 1:1 客製化設計療程 打造協調精緻的小臉。"}}},{"k":"왜 얼굴지방분해주사인가요?","t":{"en":{"e":"Why face fat dissolve injection?"},"zh-TW":{"e":"為什麼選擇臉部溶脂注射？"}}},{"k":"얼굴지방분해주사인가요?","t":{"en":{"e":"Face fat dissolve injection?"},"zh-TW":{"e":"為什麼選擇臉部溶脂注射？"}}},{"k":"1:1 맞춤 디자인 설계","t":{"en":{"e":"1:1 customized design planning"},"zh-TW":{"e":"1:1 客製化設計規劃"}}},{"k":"약 10분 내외 시술","t":{"en":{"e":"Treatment in approx. 10 minutes"},"zh-TW":{"e":"約10分鐘左右療程"}}},{"k":"멍·붓기 적음","t":{"en":{"e":"Minimal bruising and swelling"},"zh-TW":{"e":"將瘀青與腫脹 降至最低"}}},{"k":"지방세포 파괴 후 자연스럽게 체외 배출","t":{"en":{"e":"Fat cells are destroyed and naturally excreted from the body"},"zh-TW":{"e":"脂肪細胞破壞後 自然排出體外"}}},{"k":"#맞춤디자인","t":{"en":{"e":"#CustomDesign"},"zh-TW":{"e":"#客製化設計"}}},{"k":"#섬세한시술","t":{"en":{"e":"#DelicateTreatment"},"zh-TW":{"e":"#細緻療程"}}},{"k":"#슬림한얼굴라인","t":{"en":{"e":"#SlimFaceLine"},"zh-TW":{"e":"#纖細臉部線條"}}},{"k":"작아진 얼굴, 자연스러운 변화로 시작하세요","t":{"en":{"e":"A smaller face — start with natural change"},"zh-TW":{"e":"更小的臉型，從自然的改變開始"}}},{"k":"3D 볼륨필러 5cc","t":{"en":{"e":"3D Volume Filler 5cc"},"zh-TW":{"e":"3D 立體填充 5cc"}},"c":"필러"},{"k":"3D 볼륨필러","t":{"en":{"e":"3D Volume Filler"},"zh-TW":{"e":"3D 立體填充"}},"c":"필러"},{"k":"3D 볼륨 필러","t":{"en":{"e":"3D Volume Filler"},"zh-TW":{"e":"3D 立體填充"}},"c":"필러"},{"k":"3D VOLUME FILLER","t":{"en":{"e":"3D VOLUME FILLER"},"zh-TW":{"e":"3D VOLUME FILLER"}}},{"k":"미미썸 3D 볼륨 필러","t":{"en":{"e":"MIMISOME 3D Volume Filler"},"zh-TW":{"e":"MIMISOME 3D 立體填充"}}},{"k":"미미썸 3D 볼륨필러","t":{"en":{"e":"MIMISOME 3D Volume Filler"},"zh-TW":{"e":"MIMISOME 3D 立體填充"}}},{"k":"빠른 효과로 입체적인 볼륨 완성","t":{"en":{"e":"Fast results for three-dimensional volume"},"zh-TW":{"e":"快速見效 打造立體飽滿輪廓"}}},{"k":"입체적인 윤곽","t":{"en":{"e":"Three-dimensional contour"},"zh-TW":{"e":"立體輪廓"}}},{"k":"자연스러운 볼륨","t":{"en":{"e":"Natural volume"},"zh-TW":{"e":"自然飽滿"}}},{"k":"균형 잡힌 라인","t":{"en":{"e":"Balanced lines"},"zh-TW":{"e":"平衡線條"}}},{"k":"추천 대상","t":{"en":{"e":"Recommended for"},"zh-TW":{"e":"推薦對象"}}},{"k":"꺼진 부위 볼륨을 자연스럽게 채우고 싶은 분","t":{"en":{"e":"Those who want to naturally fill sunken areas"},"zh-TW":{"e":"希望自然填補凹陷部位者"}}},{"k":"얼굴 윤곽과 입체감을 함께 개선하고 싶은 분","t":{"en":{"e":"Those who want to improve both facial contour and dimension"},"zh-TW":{"e":"想同時改善臉部輪廓與比例者"}}},{"k":"자연스럽게 얼굴 균형을 맞추고 싶은 분","t":{"en":{"e":"Those who want to naturally balance their face"},"zh-TW":{"e":"希望自然調整臉部平衡者"}}},{"k":"안전성","t":{"en":{"e":"Safety"},"zh-TW":{"e":"安全性"}}},{"k":"지속력","t":{"en":{"e":"Longevity / Durability"},"zh-TW":{"e":"持久性"}}},{"k":"부위별 구조를 고려한 3D 볼륨 설계로 꺼진 부위는 채우고, 과한 볼륨은 줄여 자연스럽고 입체적인 얼굴라인을 완성합니다.","t":{"en":{"e":"With 3D volume design considering each area's structure, sunken areas are filled and excessive volume is reduced to complete a natural, three-dimensional face line."},"zh-TW":{"e":"依照部位結構設計的3D填充方案 填補凹陷部位，減少過度填充 打造自然立體的臉部線條。"}}},{"k":"시술시간","t":{"en":{"e":"Treatment time"},"zh-TW":{"e":"療程時間"}}},{"k":"일상생활","t":{"en":{"e":"Daily life"},"zh-TW":{"e":"日常生活"}}},{"k":"시술 후 빠르게 가능","t":{"en":{"e":"Resume quickly after treatment"},"zh-TW":{"e":"術後可快速恢復"}}},{"k":"통증정도","t":{"en":{"e":"Pain level"},"zh-TW":{"e":"疼痛程度"}}},{"k":"미미썸의원만의 특별함","t":{"en":{"e":"What makes MIMISOME special"},"zh-TW":{"e":"MIMISOME的獨特優勢"}}},{"k":"사전상담, 섬세한 검진을 통해 개인 상태에 맞는 맞춤형 시술 진행","t":{"en":{"e":"Through pre-consultation and detailed examination, treatment is tailored to your individual condition"},"zh-TW":{"e":"透過術前諮詢與細緻檢查 規劃符合個人狀況的療程"}}},{"k":"풍부한 경험을 바탕으로 편안하고 정교한 시술 제공","t":{"en":{"e":"Comfortable and precise treatment based on rich experience"},"zh-TW":{"e":"憑藉豐富的臨床經驗 提供舒適且精細的療程"}}},{"k":"개인의 신체특성을 고려한 자연스러운 디자인으로 만족도 향상","t":{"en":{"e":"Improved satisfaction with natural design considering individual body characteristics"},"zh-TW":{"e":"考量個人身體特性 打造自然協調的設計"}}},{"k":"108,900원","t":{"en":{"e":"108,900 KRW"},"zh-TW":{"e":"108,900 韓元"}}},{"k":"330,000원","t":{"en":{"e":"330,000 KRW"},"zh-TW":{"e":"330,000 KRW"}}},{"k":"33만","t":{"en":{"e":"330,000 KRW"},"zh-TW":{"e":"33萬韓元"}}},{"k":"*부가세 포함","t":{"en":{"e":"*VAT included"},"zh-TW":{"e":"※含稅"}}},{"k":"부가세 포함","t":{"en":{"e":"VAT included"},"zh-TW":{"e":"含稅"}}}],"glossary":{"하이톡스":{"en":"Hutox","ja":"ハイトックス","zh":"HiTox","zh-TW":"HiTox","th":"Hitox","vi":"Hyttox"},"코어톡스":{"en":"Coretox","ja":"コアトックス","zh":"Coretox","zh-TW":"Coretox","th":"Coretox","vi":"Coretox"},"이노톡스":{"en":"Innotox","ja":"イノトックス","zh":"Innotox","zh-TW":"Innotox","th":"Innotox","vi":"Innotox"},"제오민":{"en":"Xeomin","ja":"ゼオミン","zh":"Xeomin","zh-TW":"Xeomin","th":"Xeomin","vi":"Xeomin"},"디스포트":{"en":"Dysport","ja":"ディスポート","zh":"Dysport","zh-TW":"Dysport","th":"Dysport","vi":"Dysport"},"제테마":{"en":"Jetema","ja":"ジェテマ","zh":"Jetema","zh-TW":"Jetema","th":"Jetema","vi":"Jetema"},"보톡스":{"en":"Botox","ja":"ボトックス","zh":"肉毒素","zh-TW":"肉毒素","th":"โบท็อกซ์","vi":"Botox"},"필러":{"en":"Filler","ja":"フィラー","zh":"填充剂","zh-TW":"填充劑","th":"ฟิลเลอร์","vi":"Filler"},"리프팅":{"en":"Lifting","ja":"リフトアップ","zh":"提升","zh-TW":"提升","th":"ยกกระชับ","vi":"Nâng cơ"},"스킨부스터":{"en":"Skin Booster","ja":"スキンブースター","zh":"皮肤增强剂","zh-TW":"皮膚增強劑","th":"สกินบูสเตอร์","vi":"Skin Booster"},"실리프팅":{"en":"Thread Lifting","ja":"糸リフト","zh":"线雕","zh-TW":"線雕","th":"ร้อยไหม","vi":"Căng chỉ"},"윤곽주사":{"en":"Face Contouring Injection","ja":"輪郭注射","zh":"轮廓针","zh-TW":"輪廓針","th":"ฉีดสลายไขมัน","vi":"Tiêm thon gọn"},"스킨보톡스":{"en":"Skin Botox","ja":"スキンボトックス","zh":"皮肤肉毒素","zh-TW":"皮膚肉毒素","th":"สกินโบท็อกซ์","vi":"Skin Botox"},"사각턱":{"en":"Jawline / Masseter","ja":"エラ","zh":"方下巴","zh-TW":"方下巴","th":"กราม","vi":"Hàm vuông"},"쥬베룩":{"en":"Juvelook","ja":"ジュベルック","zh":"Juvelook","zh-TW":"Juvelook","th":"Juvelook","vi":"Juvelook"},"리쥬란":{"en":"Rejuran","ja":"リジュラン","zh":"Rejuran","zh-TW":"Rejuran","th":"Rejuran","vi":"Rejuran"},"스킨부스터 리투오":{"en":"Rituo Skin Booster","ja":"リトゥオスキンブースター","zh":"Rituo 皮肤增强剂","zh-TW":"Rituo 皮膚增強劑","th":"Rituo Skin Booster","vi":"Rituo Skin Booster"},"아쿠아필":{"en":"Aqua Peel","ja":"アクアピール","zh":"水光焕肤","zh-TW":"水光煥膚","th":"Aqua Peel","vi":"Aqua Peel"},"아포지":{"en":"Apogee","ja":"アポジー","zh":"Apogee","zh-TW":"Apogee","th":"Apogee","vi":"Apogee"},"피코토닝":{"en":"Pico Toning","ja":"ピコトーニング","zh":"皮秒调理","zh-TW":"皮秒調理","th":"Pico Toning","vi":"Pico Toning"},"이마":{"en":"Forehead","ja":"おでこ","zh":"额头","zh-TW":"額頭","th":"หน้าผาก","vi":"trán"},"미간":{"en":"Frown Lines","ja":"眉間","zh":"眉间","zh-TW":"眉間","th":"ระหว่างคิ้ว","vi":"giữa chân mày"},"눈가":{"en":"Crow's Feet","ja":"目元","zh":"眼周","zh-TW":"眼周","th":"หางตา","vi":"đuôi mắt"},"콧대":{"en":"Nasal Bridge","ja":"鼻柱","zh":"鼻梁","zh-TW":"鼻梁","th":"สันจมูก","vi":"sống mũi"},"콧등":{"en":"Bunny Lines","ja":"鼻背","zh":"鼻背","zh-TW":"鼻背","th":"ดั้งจมูก","vi":"lưng mũi"},"자갈턱":{"en":"Chin","ja":"顎の梅干しじわ","zh":"酒窝下巴","zh-TW":"酒窩下巴","th":"คาง","vi":"cằm đá cuội"},"국산":{"en":"Korean","ja":"韓国製","zh":"韩国产","zh-TW":"韓國產","th":"เกาหลี","vi":"nội địa"},"수입":{"en":"Imported","ja":"海外製","zh":"进口","zh-TW":"進口","th":"นำเข้า","vi":"nhập khẩu"},"1부위":{"en":"1 Area","ja":"1部位","zh":"1部位","zh-TW":"1部位","th":"1 บริเวณ","vi":"1 vùng"},"풀페이스":{"en":"Full Face","ja":"フルフェイス","zh":"全脸","zh-TW":"全臉","th":"ทั่วหน้า","vi":"Toàn mặt"},"1회":{"en":"1 Session","ja":"1回","zh":"1次","zh-TW":"1次","th":"1 ครั้ง","vi":"1 lần"},"체험가":{"en":"Trial Price","ja":"体験価格","zh":"体验价","zh-TW":"體驗價","th":"ราคาทดลอง","vi":"Giá trải nghiệm"},"첫방문":{"en":"First Visit","ja":"初回来院","zh":"首次到访","zh-TW":"首次到訪","th":"เยี่ยมชมครั้งแรก","vi":"Lần đầu đến"},"남자":{"en":"Male","ja":"男性","zh":"男士","zh-TW":"男士","th":"ผู้ชาย","vi":"Nam"},"여자":{"en":"Female","ja":"女性","zh":"女士","zh-TW":"女士","th":"ผู้หญิง","vi":"Nữ"},"택1":{"en":"Choose 1","ja":"1選択","zh":"选1","zh-TW":"選1","th":"เลือก 1","vi":"Chọn 1"},"미미썸":{"en":"Mimisome","ja":"ミミソム","zh":"美美썸","zh-TW":"MIMISOME","th":"Mimisome","vi":"Mimisome"},"미미썸 한정":{"en":"Mimisome Limited","ja":"ミミソム限定","zh":"MIMISOME 限定","zh-TW":"MIMISOME 限定","th":"Mimisome Limited","vi":"Ưu đãi Mimisome"},"한정":{"en":"Limited","ja":"限定","zh":"限定","zh-TW":"限定","th":"Limited","vi":"Giới hạn"},"VAT포함":{"en":"VAT included","ja":"VAT込み","zh":"含增值税","zh-TW":"含增值稅","th":"รวม VAT","vi":"Đã bao gồm VAT"},"메조주사":{"en":"Mesotherapy Injection","ja":"メソセラピー注射","zh":"中胚层疗法注射","zh-TW":"中胚層療法注射","th":"Mesotherapy","vi":"Tiêm mesotherapy"},"메조테라피":{"en":"Mesotherapy","ja":"メソセラピー","zh":"中胚层疗法","zh-TW":"中胚層療法","th":"Mesotherapy","vi":"Mesotherapy"},"두피":{"en":"scalp","ja":"頭皮","zh":"头皮","zh-TW":"頭皮","th":"หนังศีรษะ","vi":"da đầu"},"모근":{"en":"hair root","ja":"毛根","zh":"毛根","zh-TW":"毛根","th":"รากผม","vi":"chân tóc"},"모낭":{"en":"hair follicle","ja":"毛包","zh":"毛囊","zh-TW":"毛囊","th":"รูขุมขน","vi":"nang tóc"},"탈모":{"en":"hair loss","ja":"脱毛/発毛","zh":"脱发","zh-TW":"落髮","th":"ผมร่วง","vi":"rụng tóc"},"탈모치료":{"en":"Hair Loss Treatment","ja":"発毛治療","zh":"脱发治疗","zh-TW":"落髮治療","th":"รักษาผมร่วง","vi":"Điều trị rụng tóc"},"약물요법":{"en":"medication therapy","ja":"薬物療法","zh":"药物疗法","zh-TW":"藥物療法","th":"การรักษาด้วยยา","vi":"liệu pháp dùng thuốc"},"약물처방":{"en":"Medication Prescription","ja":"薬処方","zh":"药物处方","zh-TW":"藥物處方","th":"ยาที่สั่งโดยแพทย์","vi":"kê đơn thuốc"},"광에너지":{"en":"light energy","ja":"光エネルギー","zh":"光能","zh-TW":"光能","th":"พลังงานแสง","vi":"năng lượng ánh sáng"},"테라피":{"en":"therapy","ja":"療法","zh":"疗法","zh-TW":"療法","th":"บำบัด","vi":"liệu pháp"},"시술":{"en":"procedure","ja":"施術","zh":"治疗","zh-TW":"療程","th":"การรักษา","vi":"thủ thuật"},"마취":{"en":"anesthesia","ja":"麻酔","zh":"麻醉","zh-TW":"麻醉","th":"ยาชา","vi":"gây tê"},"연고 마취":{"en":"topical anesthesia","ja":"麻酔クリーム","zh":"表面麻醉","zh-TW":"表面麻醉","th":"ยาชาทา","vi":"gây tê bề mặt"},"일상생활":{"en":"daily life","ja":"日常生活","zh":"日常生活","zh-TW":"日常生活","th":"ชีวิตประจำวัน","vi":"sinh hoạt hàng ngày"},"복귀":{"en":"return","ja":"復帰","zh":"恢复","zh-TW":"恢復","th":"กลับสู่","vi":"trở lại"},"추천 주기":{"en":"recommended cycle","ja":"推奨周期","zh":"推荐周期","zh-TW":"建議週期","th":"ความถี่แนะนำ","vi":"tần suất khuyến nghị"},"맞춤 상담":{"en":"customized consultation","ja":"個別カウンセリング","zh":"定制咨询","zh-TW":"客製化諮詢","th":"ปรึกษาส่วนบุคคล","vi":"tư vấn cá nhân hoá"},"섬유아세포":{"en":"fibroblast","ja":"線維芽細胞","zh":"纤维母细胞","zh-TW":"纖維母細胞","th":"ไฟโบรบลาสต์","vi":"tế bào sợi"},"재생":{"en":"regeneration","ja":"再生","zh":"再生","zh-TW":"再生","th":"ฟื้นฟู","vi":"tái tạo"},"산후":{"en":"postpartum","ja":"産後","zh":"产后","zh-TW":"產後","th":"หลังคลอด","vi":"sau sinh"},"은밀 부위":{"en":"intimate areas","ja":"デリケートゾーン","zh":"私密部位","zh-TW":"私密部位","th":"จุดซ่อนเร้น","vi":"vùng kín"},"여성 은밀 부위":{"en":"women's intimate areas","ja":"女性デリケートゾーン","zh":"女性私密部位","zh-TW":"女性私密部位","th":"จุดซ่อนเร้นสำหรับผู้หญิง","vi":"vùng kín của phụ nữ"},"질":{"en":"vaginal region","ja":"膣","zh":"阴道","zh-TW":"陰道","th":"บริเวณช่องคลอด","vi":"âm đạo"},"탄력":{"en":"elasticity","ja":"弾力","zh":"弹性","zh-TW":"彈性","th":"ความยืดหยุ่น","vi":"độ đàn hồi"},"볼륨":{"en":"volume","ja":"ボリューム","zh":"丰盈","zh-TW":"豐盈","th":"วอลลุ่ม","vi":"thể tích"},"볼륨감":{"en":"volume","ja":"ボリューム感","zh":"饱满感","zh-TW":"豐盈感","th":"ความฟูดูเป็นธรรมชาติ","vi":"cảm giác đầy đặn"},"자신감":{"en":"confidence","ja":"自信","zh":"自信","zh-TW":"自信","th":"ความมั่นใจ","vi":"sự tự tin"},"출산":{"en":"childbirth","ja":"出産","zh":"分娩","zh-TW":"生產","th":"การคลอด","vi":"sinh nở"},"노화":{"en":"aging","ja":"加齢","zh":"老化","zh-TW":"老化","th":"อายุที่เพิ่มขึ้น","vi":"lão hóa"},"임상 경험":{"en":"clinical experience","ja":"臨床経験","zh":"临床经验","zh-TW":"臨床經驗","th":"ประสบการณ์ทางคลินิก","vi":"kinh nghiệm lâm sàng"},"통증":{"en":"pain","ja":"痛み","zh":"疼痛","zh-TW":"疼痛","th":"ความเจ็บ","vi":"đau"},"회복":{"en":"recovery","ja":"回復","zh":"恢复","zh-TW":"恢復","th":"การฟื้นตัว","vi":"phục hồi"},"1:1 맞춤":{"en":"1:1 personalized","ja":"1:1 オーダーメイド","zh":"1:1 定制","zh-TW":"1:1 客製化","th":"ตัวต่อตัว","vi":"1:1 cá nhân hóa"},"지속성":{"en":"longevity","ja":"持続性","zh":"持久性","zh-TW":"持久性","th":"ความคงทน","vi":"tính bền lâu"},"안전성":{"en":"safety","ja":"安全性","zh":"安全性","zh-TW":"安全性","th":"ความปลอดภัย","vi":"tính an toàn"},"자연스러운":{"en":"natural","ja":"自然な","zh":"自然","zh-TW":"自然","th":"เป็นธรรมชาติ","vi":"tự nhiên"},"필러 성분":{"en":"filler ingredients","ja":"フィラー成分","zh":"填充成分","zh-TW":"填充成分","th":"ส่วนผสมฟิลเลอร์","vi":"thành phần filler"},"히알루론산":{"en":"hyaluronic acid","ja":"ヒアルロン酸","zh":"玻尿酸","zh-TW":"玻尿酸","th":"กรดไฮยาลูโรนิก","vi":"axit hyaluronic"},"인티메이트":{"en":"intimate","ja":"インティメート","zh":"私密","zh-TW":"親密","th":"อินทิเมท","vi":"vùng kín"},"여성 인티메이트":{"en":"women's intimate area","ja":"女性インティメート","zh":"女性私密","zh-TW":"女性親密部位","th":"อินทิเมทเฉพาะผู้หญิง","vi":"vùng kín nữ giới"},"성감":{"en":"sensitivity","ja":"感度","zh":"敏感度","zh-TW":"敏感度","th":"ความรู้สึก","vi":"độ nhạy cảm"},"지속력":{"en":"longevity","ja":"持続力","zh":"持久力","zh-TW":"持久力","th":"ความคงทน","vi":"tính bền lâu"},"맞춤상담":{"en":"personalized consultation","ja":"オーダーメイドカウンセリング","zh":"定制咨询","zh-TW":"客製化諮詢","th":"ปรึกษาเฉพาะบุคคล","vi":"tư vấn cá nhân hóa"},"심미성":{"en":"aesthetics","ja":"審美性","zh":"审美","zh-TW":"美感","th":"ความสวยงาม","vi":"thẩm mỹ"},"기능성":{"en":"functionality","ja":"機能性","zh":"功能性","zh-TW":"功能性","th":"การใช้งาน","vi":"tính năng"},"의료진":{"en":"medical staff","ja":"医療スタッフ","zh":"医疗团队","zh-TW":"醫療團隊","th":"ทีมแพทย์","vi":"đội ngũ y tế"},"노하우":{"en":"expertise","ja":"ノウハウ","zh":"经验技术","zh-TW":"經驗技術","th":"ความเชี่ยวชาญ","vi":"bí quyết"},"미미썸의원":{"en":"MIMISOME Clinic","ja":"ミミソム医院","zh":"MIMISOME诊所","zh-TW":"MIMISOME診所","th":"คลินิก MIMISOME","vi":"Phòng khám MIMISOME"},"엑소좀":{"en":"exosome","ja":"エクソソーム","zh":"外泌体","zh-TW":"外泌體","th":"เอ็กโซโซม","vi":"exosome"},"ASCE+":{"en":"ASCE+","ja":"ASCE+","zh":"ASCE+","zh-TW":"ASCE+","th":"ASCE+","vi":"ASCE+"},"콜라겐":{"en":"collagen","ja":"コラーゲン","zh":"胶原蛋白","zh-TW":"膠原蛋白","th":"คอลลาเจน","vi":"collagen"},"엘라스틴":{"en":"elastin","ja":"エラスチン","zh":"弹性蛋白","zh-TW":"彈力蛋白","th":"อีลาสติน","vi":"elastin"},"모공":{"en":"pore","ja":"毛穴","zh":"毛孔","zh-TW":"毛孔","th":"รูขุมขน","vi":"lỗ chân lông"},"모공 축소":{"en":"pore care","ja":"毛穴ケア","zh":"毛孔缩小","zh-TW":"毛孔護理","th":"ดูแลรูขุมขน","vi":"thu nhỏ lỗ chân lông"},"모공지우개":{"en":"Pore Care","ja":"毛穴ケア","zh":"毛孔护理","zh-TW":"毛孔護理","th":"ลบเลือนรูขุมขน","vi":"xóa lỗ chân lông"},"피부결":{"en":"skin texture","ja":"肌キメ","zh":"肤质","zh-TW":"膚質","th":"ผิวสัมผัส","vi":"kết cấu da"},"유수분":{"en":"moisture and oil balance","ja":"水分油分","zh":"水油平衡","zh-TW":"水油平衡","th":"ความชุ่มชื้นและน้ำมัน","vi":"cân bằng dầu và nước"},"유수분 밸런스":{"en":"Hydration Balance","ja":"水分バランス","zh":"水油平衡","zh-TW":"水分平衡","th":"สมดุลความชุ่มชื้น","vi":"cân bằng độ ẩm"},"트러블":{"en":"skin problems / acne","ja":"肌トラブル","zh":"肌肤问题","zh-TW":"肌膚問題","th":"ปัญหาผิว","vi":"vấn đề da"},"침습적":{"en":"invasive","ja":"侵襲的","zh":"侵入式","zh-TW":"侵入式","th":"แบบรุกราน","vi":"xâm lấn"},"비침습":{"en":"non-invasive","ja":"非侵襲","zh":"非侵入式","zh-TW":"非侵入式","th":"ไม่รุกราน","vi":"không xâm lấn"},"근본":{"en":"fundamental / root cause","ja":"根本","zh":"根本","zh-TW":"根本","th":"พื้นฐาน","vi":"căn nguyên"},"광택":{"en":"glow","ja":"ツヤ","zh":"光泽","zh-TW":"光澤","th":"ความเงา","vi":"độ bóng"},"팅커벨":{"en":"Tinker Bell","ja":"ティンカーベル","zh":"小精灵","zh-TW":"精靈","th":"ทิงเกอร์เบลล์","vi":"Tinker Bell"},"귀":{"en":"ear","ja":"耳","zh":"耳朵","zh-TW":"耳朵","th":"หู","vi":"tai"},"귓볼":{"en":"earlobe","ja":"耳たぶ","zh":"耳垂","zh-TW":"耳垂","th":"ติ่งหู","vi":"dái tai"},"귀필러":{"en":"ear filler","ja":"耳フィラー","zh":"耳部填充","zh-TW":"精靈耳玻尿酸","th":"ฟิลเลอร์หู","vi":"filler tai"},"소얼굴":{"en":"small face / face slimming","ja":"小顔","zh":"小脸","zh-TW":"小顏","th":"หน้าเล็ก","vi":"mặt nhỏ"},"얼굴 비율":{"en":"facial proportions","ja":"顔のバランス","zh":"脸部比例","zh-TW":"臉部比例","th":"สัดส่วนใบหน้า","vi":"tỉ lệ khuôn mặt"},"이목구비":{"en":"facial features","ja":"顔のパーツ","zh":"五官","zh-TW":"五官","th":"ใบหน้า","vi":"ngũ quan"},"비대칭":{"en":"asymmetrical","ja":"非対称","zh":"不对称","zh-TW":"不對稱","th":"ไม่สมมาตร","vi":"không đối xứng"},"이상적인":{"en":"ideal","ja":"理想的な","zh":"理想的","zh-TW":"理想的","th":"ที่เหมาะสมที่สุด","vi":"lý tưởng"},"비율":{"en":"proportions","ja":"比率","zh":"比例","zh-TW":"比例","th":"สัดส่วน","vi":"tỉ lệ"},"균형":{"en":"balance","ja":"バランス","zh":"平衡","zh-TW":"平衡","th":"สมดุล","vi":"cân bằng"},"조화로운":{"en":"harmonious / balanced","ja":"調和のとれた","zh":"协调","zh-TW":"協調","th":"กลมกลืน","vi":"hài hòa"},"비결":{"en":"secret / tip","ja":"秘訣","zh":"秘诀","zh-TW":"秘訣","th":"เคล็ดลับ","vi":"bí quyết"},"얇은":{"en":"thin","ja":"薄い","zh":"薄","zh-TW":"薄","th":"บาง","vi":"mỏng"},"정면":{"en":"front view","ja":"正面","zh":"正面","zh-TW":"正面","th":"ด้านหน้า","vi":"mặt trước"},"써마지":{"en":"Thermage","ja":"サーマクール","zh":"Thermage","zh-TW":"鳳凰電波","th":"เทอร์มาจ","vi":"Thermage"},"써마지FLX":{"en":"Thermage FLX","ja":"サーマクールFLX","zh":"Thermage FLX","zh-TW":"第四代鳳凰電波","th":"เทอร์มาจ FLX","vi":"Thermage FLX"},"써마지 FLX":{"en":"Thermage FLX","ja":"サーマクールFLX","zh":"Thermage FLX","zh-TW":"第四代鳳凰電波","th":"เทอร์มาจ FLX","vi":"Thermage FLX"},"FLX":{"en":"FLX","ja":"FLX","zh":"FLX","zh-TW":"FLX","th":"FLX","vi":"FLX"},"고주파":{"en":"radiofrequency / RF","ja":"高周波","zh":"高频","zh-TW":"電波","th":"คลื่นวิทยุ","vi":"sóng RF"},"고주파 에너지":{"en":"radiofrequency energy","ja":"高周波エネルギー","zh":"射频能量","zh-TW":"電波能量","th":"พลังงานคลื่นวิทยุ","vi":"năng lượng sóng RF"},"잔주름":{"en":"fine lines","ja":"小ジワ","zh":"细纹","zh-TW":"細紋","th":"ริ้วรอยเล็กๆ","vi":"nếp nhăn nhỏ"},"윤곽":{"en":"contour","ja":"輪郭","zh":"轮廓","zh-TW":"輪廓","th":"โครงหน้า","vi":"đường viền khuôn mặt"},"타이트닝":{"en":"tightening","ja":"タイトニング","zh":"紧致","zh-TW":"緊實","th":"กระชับ","vi":"săn chắc"},"스킨 타이트닝":{"en":"skin tightening","ja":"スキンタイトニング","zh":"皮肤紧致","zh-TW":"肌膚緊實","th":"กระชับผิว","vi":"săn chắc da"},"300샷":{"en":"300 Shots","ja":"300ショット","zh":"300发","zh-TW":"300發","th":"300 ช็อต","vi":"300 shot"},"샷":{"en":"shot","ja":"ショット","zh":"发","zh-TW":"發","th":"ช็อต","vi":"shot"},"처짐":{"en":"sagging","ja":"たるみ","zh":"松弛","zh-TW":"鬆弛","th":"หย่อนคล้อย","vi":"chảy xệ"},"처진":{"en":"sagging","ja":"たるんだ","zh":"松弛的","zh-TW":"鬆弛的","th":"หย่อนคล้อย","vi":"chảy xệ"},"턱선":{"en":"jawline","ja":"フェイスライン","zh":"下颌线","zh-TW":"下顎線","th":"กรอบหน้า","vi":"đường viền hàm"},"얼굴라인":{"en":"facial contours","ja":"フェイスライン","zh":"脸部线条","zh-TW":"臉部輪廓","th":"โครงหน้า","vi":"đường nét khuôn mặt"},"입가":{"en":"around the mouth","ja":"口元","zh":"嘴角","zh-TW":"嘴角","th":"รอบปาก","vi":"vùng quanh miệng"},"주름":{"en":"wrinkle","ja":"シワ","zh":"皱纹","zh-TW":"皺紋","th":"ริ้วรอย","vi":"nếp nhăn"},"진동":{"en":"vibration","ja":"振動","zh":"震动","zh-TW":"震動","th":"การสั่นสะเทือน","vi":"rung"},"전방향":{"en":"multidirectional","ja":"全方向","zh":"全方位","zh-TW":"全方位","th":"รอบทิศทาง","vi":"đa hướng"},"맞춤":{"en":"customized","ja":"オーダーメイド","zh":"定制","zh-TW":"客製化","th":"ปรับเฉพาะบุคคล","vi":"cá nhân hoá"},"광고용 이미지":{"en":"advertising image","ja":"広告用イメージ","zh":"广告形象","zh-TW":"廣告形象","th":"ภาพโฆษณา","vi":"hình ảnh quảng cáo"},"포텐자":{"en":"Potenza","ja":"ポテンツァ","zh":"Potenza","zh-TW":"Potenza","th":"โพเทนซ่า","vi":"Potenza"},"포텐자 레이저":{"en":"Potenza Laser","ja":"ポテンツァレーザー","zh":"Potenza激光","zh-TW":"Potenza雷射","th":"โพเทนซ่าเลเซอร์","vi":"Potenza Laser"},"레이저":{"en":"laser","ja":"レーザー","zh":"激光","zh-TW":"雷射","th":"เลเซอร์","vi":"laser"},"모공축소":{"en":"pore reduction","ja":"毛穴縮小","zh":"缩小毛孔","zh-TW":"縮小毛孔","th":"ลดรูขุมขน","vi":"thu nhỏ lỗ chân lông"},"흉터":{"en":"scar","ja":"傷跡","zh":"疤痕","zh-TW":"疤痕","th":"แผลเป็น","vi":"sẹo"},"흉터치료":{"en":"scar treatment","ja":"傷跡治療","zh":"疤痕治疗","zh-TW":"疤痕治療","th":"รักษาแผลเป็น","vi":"điều trị sẹo"},"여드름":{"en":"acne","ja":"ニキビ","zh":"痤疮","zh-TW":"痘痘","th":"สิว","vi":"mụn trứng cá"},"여드름 흉터":{"en":"acne scar","ja":"ニキビ跡","zh":"痘印 / 痘疤","zh-TW":"痘疤","th":"แผลสิว","vi":"sẹo mụn"},"탄력개선":{"en":"elasticity improvement","ja":"ハリ改善","zh":"弹性改善","zh-TW":"彈力改善","th":"เพิ่มความกระชับ","vi":"cải thiện độ đàn hồi"},"깐달걀피부":{"en":"glass skin","ja":"ゆで卵肌","zh":"剥壳鸡蛋肌","zh-TW":"剝殼雞蛋肌","th":"ผิวไข่ต้ม","vi":"da trứng gà bóc vỏ"},"자신감UP":{"en":"confidence UP","ja":"自信アップ","zh":"自信UP","zh-TW":"自信UP","th":"เพิ่มความมั่นใจ","vi":"tự tin UP"},"피부층":{"en":"skin layer","ja":"皮膚層","zh":"皮肤层","zh-TW":"皮膚層","th":"ชั้นผิว","vi":"lớp da"},"깊이":{"en":"depth","ja":"深さ","zh":"深度","zh-TW":"深度","th":"ความลึก","vi":"độ sâu"},"종류":{"en":"type","ja":"種類","zh":"种类","zh-TW":"種類","th":"ชนิด","vi":"loại"},"에너지":{"en":"energy","ja":"エネルギー","zh":"能量","zh-TW":"能量","th":"พลังงาน","vi":"năng lượng"},"시술횟수":{"en":"number of sessions","ja":"施術回数","zh":"治疗次数","zh-TW":"療程次數","th":"จำนวนครั้งที่ทำ","vi":"số lần điều trị"},"키메 / 피부결":{"en":"skin texture","ja":"キメ","zh":"肤质","zh-TW":"膚質","th":"ผิวสัมผัส","vi":"kết cấu da"},"RF":{"en":"RF","ja":"RF / 高周波","zh":"射频","zh-TW":"電波","th":"RF","vi":"sóng RF"},"카운셀링":{"en":"consultation","ja":"カウンセリング","zh":"咨询","zh-TW":"諮詢","th":"คำปรึกษา","vi":"tư vấn"},"카운셀링 포함":{"en":"consultation included","ja":"カウンセリング込み","zh":"含咨询","zh-TW":"含諮詢","th":"รวมค่าปรึกษา","vi":"đã bao gồm tư vấn"},"리프트업":{"en":"lift-up","ja":"リフトアップ","zh":"提升","zh-TW":"提拉","th":"ยกกระชับ","vi":"nâng cơ"},"예민":{"en":"sensitive","ja":"敏感","zh":"敏感","zh-TW":"敏感","th":"แพ้ง่าย","vi":"nhạy cảm"},"리쥬란HB":{"en":"REJURAN HB","ja":"リジュランHB","zh":"丽珠兰 HB","zh-TW":"麗珠蘭 HB","th":"REJURAN HB","vi":"Rejuran HB"},"리쥬란HB PLUS":{"en":"REJURAN HB PLUS","ja":"リジュラン HB PLUS","zh":"丽珠兰 HB PLUS","zh-TW":"麗珠蘭 HB PLUS","th":"REJURAN HB PLUS","vi":"Rejuran HB PLUS"},"리쥬란 힐러":{"en":"Rejuran Healer","ja":"リジュラン ヒーラー / リジュラン成分","zh":"丽珠兰成分","zh-TW":"麗珠蘭成分","th":"สารรีจูรัน","vi":"Rejuran Healer"},"HB PLUS":{"en":"HB PLUS","ja":"HB PLUS","zh":"HB PLUS","zh-TW":"HB PLUS","th":"HB PLUS","vi":"HB PLUS"},"안티에이징":{"en":"anti-aging","ja":"アンチエイジング","zh":"抗衰老","zh-TW":"抗老化","th":"การชะลอวัย","vi":"chống lão hóa"},"연어주사":{"en":"salmon injection","ja":"サーモン注射","zh":"鲑鱼针","zh-TW":"鮭魚針","th":"เข็มแซลมอน","vi":"tiêm cá hồi"},"연어":{"en":"salmon","ja":"サーモン","zh":"鲑鱼","zh-TW":"鮭魚","th":"แซลมอน","vi":"cá hồi"},"폴리뉴클레오타이드":{"en":"polynucleotide / PN","ja":"ポリヌクレオチド","zh":"多核苷酸","zh-TW":"多核苷酸","th":"โพลีนิวคลีโอไทด์","vi":"polynucleotide"},"PN성분":{"en":"PN ingredient","ja":"PN成分","zh":"PN成分","zh-TW":"PN成分","th":"สาร PN","vi":"thành phần PN"},"리도카인":{"en":"lidocaine","ja":"リドカイン","zh":"利多卡因","zh-TW":"利多卡因","th":"ลิโดเคน","vi":"lidocaine"},"국소마취":{"en":"local anesthesia","ja":"局所麻酔","zh":"局部麻醉","zh-TW":"局部麻醉","th":"ยาชาเฉพาะที่","vi":"gây tê tại chỗ"},"마취성분":{"en":"anesthetic ingredient","ja":"麻酔成分","zh":"麻醉成分","zh-TW":"麻醉成分","th":"ส่วนผสมยาชา","vi":"thành phần gây tê"},"수분감":{"en":"hydration","ja":"保湿感","zh":"水润感","zh-TW":"保濕感","th":"ความชุ่มชื้น","vi":"độ ẩm"},"보습":{"en":"moisturizing / hydration","ja":"保湿","zh":"保湿","zh-TW":"保濕","th":"ให้ความชุ่มชื้น","vi":"dưỡng ẩm"},"피부재생":{"en":"skin regeneration","ja":"皮膚再生","zh":"皮肤再生","zh-TW":"皮膚再生","th":"การฟื้นฟูผิว","vi":"tái tạo da"},"탄력상승":{"en":"improved elasticity","ja":"弾力アップ","zh":"弹力提升","zh-TW":"彈力提升","th":"เพิ่มความกระชับ","vi":"tăng độ đàn hồi"},"피부결 개선":{"en":"improved skin texture","ja":"肌キメ改善","zh":"肤质改善","zh-TW":"膚質改善","th":"ผิวเรียบเนียนขึ้น","vi":"cải thiện kết cấu da"},"윤기":{"en":"radiance / glow","ja":"ツヤ","zh":"光泽","zh-TW":"光澤","th":"ความเปล่งปลั่ง","vi":"độ bóng / sáng"},"생기":{"en":"vitality","ja":"ハリ / 生気","zh":"活力","zh-TW":"活力","th":"ความสดใส","vi":"sức sống"},"푸석함":{"en":"dryness / rough","ja":"カサつき","zh":"粗糙","zh-TW":"粗糙","th":"หยาบกร้าน","vi":"thô ráp"},"건조함":{"en":"dryness","ja":"乾燥","zh":"干燥","zh-TW":"乾燥","th":"แห้ง","vi":"khô"},"분해":{"en":"break down / decompose","ja":"分解","zh":"分解","zh-TW":"分解","th":"สลายตัว","vi":"phân hủy"},"형태 유지":{"en":"maintains structure","ja":"形状を保つ","zh":"保持形态","zh-TW":"保持形態","th":"คงรูปทรง","vi":"giữ hình dạng"},"생체 적합":{"en":"biocompatible","ja":"生体適合","zh":"生物相容","zh-TW":"生物相容","th":"เข้ากันได้กับร่างกาย","vi":"tương thích sinh học"},"쥬베룩 스킨":{"en":"Juvelook Skin","ja":"ジュベルックスキン","zh":"乔雅露","zh-TW":"喬雅露","th":"Juvelook","vi":"Juvelook Skin"},"JUVELOOK":{"en":"JUVELOOK","ja":"ジュベルック","zh":"乔雅露","zh-TW":"喬雅露","th":"JUVELOOK","vi":"Juvelook"},"콜라겐 생성":{"en":"collagen production","ja":"コラーゲン生成","zh":"胶原蛋白生成","zh-TW":"膠原蛋白生成","th":"การสร้างคอลลาเจน","vi":"sản sinh collagen"},"자가콜라겐":{"en":"natural / self collagen","ja":"自己コラーゲン","zh":"自身胶原","zh-TW":"自身膠原","th":"คอลลาเจนตามธรรมชาติ","vi":"collagen tự thân"},"물광":{"en":"radiant glow / dewy","ja":"ツヤ / 水光","zh":"水光","zh-TW":"水光","th":"ผิวฉ่ำโกลว์","vi":"da căng bóng"},"물광 피부":{"en":"glass-like dewy skin","ja":"水光肌","zh":"水光肌","zh-TW":"水光肌","th":"ผิวฉ่ำวาว","vi":"da căng bóng"},"긴장":{"en":"tightening","ja":"引き締め","zh":"紧致","zh-TW":"緊緻","th":"กระชับ","vi":"săn chắc"},"PLA":{"en":"PLA (polylactic acid)","ja":"PLA","zh":"PLA","zh-TW":"PLA","th":"PLA","vi":"PLA"},"HA":{"en":"HA (hyaluronic acid)","ja":"HA","zh":"HA","zh-TW":"HA","th":"HA","vi":"HA"},"Hybrid filler":{"en":"Hybrid filler","ja":"ハイブリッドフィラー","zh":"混合填充剂","zh-TW":"混合填充劑","th":"ฟิลเลอร์ไฮบริด","vi":"filler hybrid"},"거칠다":{"en":"rough","ja":"ごわつく","zh":"粗糙","zh-TW":"粗糙","th":"หยาบ","vi":"thô ráp"},"자글한":{"en":"fine / wrinkly","ja":"細かい","zh":"细小","zh-TW":"細小","th":"เล็กๆ","vi":"nhỏ li ti"},"복합 시술":{"en":"combination treatment","ja":"総合的なケア","zh":"综合护理","zh-TW":"複合療程","th":"การดูแลแบบครบ","vi":"liệu trình kết hợp"},"AI 모델":{"en":"AI model","ja":"AIモデル","zh":"AI模型","zh-TW":"AI模型","th":"โมเดล AI","vi":"mô hình AI"},"생성형 AI":{"en":"generative AI","ja":"生成型AI","zh":"生成式AI","zh-TW":"生成式AI","th":"AI สร้างภาพ","vi":"AI tạo sinh"},"무턱":{"en":"recessed chin","ja":"後退顎","zh":"无下巴","zh-TW":"無下巴","th":"คางสั้น","vi":"cằm lẹm"},"무턱필러":{"en":"Chin Filler","ja":"無顎フィラー","zh":"无下巴填充","zh-TW":"下巴填充","th":"ฟิลเลอร์คาง","vi":"filler cằm lẹm"},"돌출입":{"en":"protruding mouth","ja":"突き出た口","zh":"突嘴","zh-TW":"下巴前凸","th":"ปากยื่น","vi":"miệng vẩu"},"이중턱":{"en":"double chin","ja":"二重あご","zh":"双下巴","zh-TW":"雙下巴","th":"เหนียงคู่","vi":"cằm đôi"},"하관":{"en":"lower face","ja":"下顔","zh":"下半脸","zh-TW":"下半臉","th":"ใบหน้าส่วนล่าง","vi":"phần dưới khuôn mặt"},"갸름주사":{"en":"slimming injection","ja":"スリミング注射","zh":"瘦脸针","zh-TW":"溶脂注射","th":"ฉีดเรียวหน้า","vi":"tiêm thon gọn"},"얼굴지방분해주사":{"en":"Face Fat Dissolve Injection","ja":"顔脂肪溶解注射","zh":"面部溶脂针","zh-TW":"臉部溶脂注射","th":"ฉีดสลายไขมันใบหน้า","vi":"tiêm tan mỡ mặt"},"지방분해주사":{"en":"fat dissolve injection","ja":"脂肪溶解注射","zh":"溶脂针","zh-TW":"溶脂注射","th":"ฉีดสลายไขมัน","vi":"tiêm tan mỡ"},"볼살":{"en":"cheek fat","ja":"頬の肉","zh":"颊脂肪","zh-TW":"臉頰脂肪","th":"ไขมันแก้ม","vi":"mỡ má"},"V라인":{"en":"V-line","ja":"Vライン","zh":"V线","zh-TW":"V型","th":"V-line","vi":"V-line"},"지방세포":{"en":"fat cell","ja":"脂肪細胞","zh":"脂肪细胞","zh-TW":"脂肪細胞","th":"เซลล์ไขมัน","vi":"tế bào mỡ"},"붓기":{"en":"swelling","ja":"腫れ","zh":"肿胀","zh-TW":"腫脹","th":"อาการบวม","vi":"sưng"},"멍":{"en":"bruising","ja":"あざ","zh":"瘀青","zh-TW":"瘀青","th":"รอยช้ำ","vi":"vết bầm"},"3D 볼륨필러":{"en":"3D Volume Filler","ja":"3Dボリュームフィラー","zh":"3D立体填充","zh-TW":"3D 立體填充","th":"3D Volume Filler","vi":"Filler 3D"},"볼륨필러":{"en":"Volume Filler","ja":"ボリュームフィラー","zh":"立体填充","zh-TW":"立體填充","th":"Volume Filler","vi":"Volume Filler"},"꺼진 부위":{"en":"sunken area","ja":"くぼんだ部分","zh":"凹陷部位","zh-TW":"凹陷部位","th":"บริเวณยุบ","vi":"vùng lõm"},"입체감":{"en":"three-dimensional / dimension","ja":"立体感","zh":"立体感","zh-TW":"立體感","th":"มิติ","vi":"hiệu ứng 3D"},"사전상담":{"en":"pre-consultation","ja":"事前カウンセリング","zh":"术前咨询","zh-TW":"術前諮詢","th":"ให้คำปรึกษาก่อนทำ","vi":"tư vấn trước thực hiện"},"맞춤시술":{"en":"customized treatment","ja":"カスタマイズ施術","zh":"定制护理","zh-TW":"客製化療程","th":"การรักษาเฉพาะบุคคล","vi":"điều trị cá nhân hóa"},"Shrink / Shurink":{"en":"Shrink / Shurink","ja":"シュリンク","zh":"修复","zh-TW":"Shurink","th":"Shrink","vi":"Shrink"},"슈링크":{"en":"Shurink","ja":"シュリンク","zh":"舒颜提拉","zh-TW":"Shurink","th":"Shurink","vi":"Shurink"}}}`;
const DATA = JSON.parse(RAW_DATA);


const LANGUAGES = [
  { code: 'en', native: 'English', flag: '🇺🇸', label: '영어', srcCol: 4 },
  { code: 'ja', native: '日本語', flag: '🇯🇵', label: '일본어', srcCol: 5 },
  { code: 'zh-TW', native: '繁體中文', flag: '🇹🇼', label: '대만어', srcCol: 6 },
  { code: 'zh', native: '简体中文', flag: '🇨🇳', label: '중국어', srcCol: 7 },
  { code: 'th', native: 'ภาษาไทย', flag: '🇹🇭', label: '태국어', srcCol: 8 },
  { code: 'vi', native: 'Tiếng Việt', flag: '🇻🇳', label: '베트남어', srcCol: 9 },
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
4. zh = Simplified Chinese; zh-TW = Traditional Chinese.
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
4. ${lang === 'zh' ? 'Use Simplified Chinese characters (中国大陆).' : ''}${lang === 'zh-TW' ? 'Use Traditional Chinese characters (繁體中文, 台灣).' : ''}
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

// ============================================================================
// 디자인 100% 보존 출력 (XLSX-as-ZIP + XML 셀 값 치환)
// XLSX는 ZIP 파일. 셀 텍스트만 sheet XML에서 직접 바꾸고 나머지 모든 바이트
// (styles.xml, 메모, 그림, 테마, 병합, drawings 등)는 원본 그대로 보존.
// 이것이 디자인이 100% 살아남는 유일한 방법.
// ============================================================================

async function inflateRaw(compressed) {
  if (compressed.length === 0) return new Uint8Array();
  const stream = new Blob([compressed]).stream();
  const decompressed = stream.pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(decompressed).arrayBuffer());
}

async function deflateRaw(raw) {
  if (raw.length === 0) return new Uint8Array();
  const stream = new Blob([raw]).stream();
  const compressed = stream.pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(compressed).arrayBuffer());
}

async function parseXlsxZip(buffer) {
  const data = new Uint8Array(buffer);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  
  // EOCD를 뒤에서부터 검색
  let eocdPos = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocdPos = i; break; }
  }
  if (eocdPos === -1) throw new Error('ZIP EOCD not found');
  
  const numEntries = view.getUint16(eocdPos + 10, true);
  const cdOffset = view.getUint32(eocdPos + 16, true);
  
  const decoder = new TextDecoder('utf-8');
  const files = new Map();
  let pos = cdOffset;
  
  for (let i = 0; i < numEntries; i++) {
    if (view.getUint32(pos, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(pos + 10, true);
    const compSize = view.getUint32(pos + 20, true);
    const uncompSize = view.getUint32(pos + 24, true);
    const nameLen = view.getUint16(pos + 28, true);
    const extraLen = view.getUint16(pos + 30, true);
    const commentLen = view.getUint16(pos + 32, true);
    const localOffset = view.getUint32(pos + 42, true);
    
    const name = decoder.decode(data.slice(pos + 46, pos + 46 + nameLen));
    
    const lhNameLen = view.getUint16(localOffset + 26, true);
    const lhExtraLen = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + lhNameLen + lhExtraLen;
    const compData = data.slice(dataStart, dataStart + compSize);
    
    files.set(name, { compData, compMethod, uncompSize });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  
  return files;
}

async function readXlsxFile(entry) {
  if (entry.compMethod === 0) return entry.compData;
  if (entry.compMethod === 8) return await inflateRaw(entry.compData);
  throw new Error('Unsupported ZIP compression: ' + entry.compMethod);
}

async function buildXlsxZip(filesMap) {
  // filesMap: Map<string, Uint8Array> 압축 안 된 원본
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  
  for (const [name, raw] of filesMap) {
    const nameBytes = encoder.encode(name);
    const compressed = await deflateRaw(raw);
    const crc = crc32(raw);
    
    const lfh = new Uint8Array(30);
    const lfhView = new DataView(lfh.buffer);
    lfhView.setUint32(0, 0x04034b50, true);
    lfhView.setUint16(4, 20, true);
    lfhView.setUint16(6, 0x0800, true);
    lfhView.setUint16(8, 8, true); // DEFLATE
    lfhView.setUint32(14, crc, true);
    lfhView.setUint32(18, compressed.length, true);
    lfhView.setUint32(22, raw.length, true);
    lfhView.setUint16(26, nameBytes.length, true);
    localParts.push(lfh, nameBytes, compressed);
    
    const cde = new Uint8Array(46);
    const cdeView = new DataView(cde.buffer);
    cdeView.setUint32(0, 0x02014b50, true);
    cdeView.setUint16(4, 20, true);
    cdeView.setUint16(6, 20, true);
    cdeView.setUint16(8, 0x0800, true);
    cdeView.setUint16(10, 8, true);
    cdeView.setUint32(16, crc, true);
    cdeView.setUint32(20, compressed.length, true);
    cdeView.setUint32(24, raw.length, true);
    cdeView.setUint16(28, nameBytes.length, true);
    cdeView.setUint32(42, offset, true);
    centralParts.push(cde, nameBytes);
    
    offset += 30 + nameBytes.length + compressed.length;
  }
  
  const centralStart = offset;
  let centralSize = 0;
  for (const p of centralParts) centralSize += p.length;
  
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(8, filesMap.size, true);
  eocdView.setUint16(10, filesMap.size, true);
  eocdView.setUint32(12, centralSize, true);
  eocdView.setUint32(16, centralStart, true);
  
  const allParts = [...localParts, ...centralParts, eocd];
  let total = 0;
  for (const p of allParts) total += p.length;
  const result = new Uint8Array(total);
  let pos = 0;
  for (const p of allParts) { result.set(p, pos); pos += p.length; }
  return result;
}

// row(0-indexed), col(0-indexed) → "A1" 형식 셀 참조
function encodeCellRef(row, col) {
  let s = '';
  let c = col;
  while (c >= 0) {
    s = String.fromCharCode(65 + (c % 26)) + s;
    c = Math.floor(c / 26) - 1;
  }
  return s + (row + 1);
}

// col(0-indexed) → "A" 컬럼 문자
function encodeColLetter(col) {
  let s = '';
  let c = col;
  while (c >= 0) {
    s = String.fromCharCode(65 + (c % 26)) + s;
    c = Math.floor(c / 26) - 1;
  }
  return s;
}

// "A" → col(0-indexed)
function decodeColLetter(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

// 데이터 영역(fromRow1Indexed 이상)의 셀들 중 col B(>=1) 이상을 오른쪽으로 1칸 시프트.
// <c r="B7"> → <c r="C7">, <c r="C7"> → <c r="D7">, ...
function shiftDataColumnsRight(sheetXml, fromRow1Indexed) {
  return sheetXml.replace(
    /(<row\s+r=")(\d+)("[^>]*>)([\s\S]*?)(<\/row>)/g,
    (match, openPre, rowNumStr, openPost, content, closeTag) => {
      const rowNum = parseInt(rowNumStr, 10);
      if (rowNum < fromRow1Indexed) return match;
      
      // 행 내 모든 r="XR" 패턴에서 X 컬럼 문자 이동
      const newContent = content.replace(
        /(\sr=")([A-Z]+)(\d+)(")/g,
        (m, pre, letter, num, post) => {
          const col = decodeColLetter(letter);
          if (col < 1) return m; // A 컬럼은 유지
          return pre + encodeColLetter(col + 1) + num + post;
        }
      );
      
      return openPre + rowNumStr + openPost + newContent + closeTag;
    }
  );
}

// <cols>에서 min/max가 fromColIndex(1-indexed) 이상인 정의 시프트 + 새 col B 삽입
function shiftColsAndInsertNewB(sheetXml, newColBWidth = 22) {
  // 기존 col 정의들 중 min >= 2 인 것을 +1 시프트
  let result = sheetXml.replace(
    /<col\s+([^>]*?)\s*\/>/g,
    (match, attrs) => {
      const minMatch = attrs.match(/\bmin="(\d+)"/);
      const maxMatch = attrs.match(/\bmax="(\d+)"/);
      if (!minMatch || !maxMatch) return match;
      const min = parseInt(minMatch[1], 10);
      const max = parseInt(maxMatch[1], 10);
      if (min < 2) return match;
      const newAttrs = attrs
        .replace(/\bmin="\d+"/, `min="${min + 1}"`)
        .replace(/\bmax="\d+"/, `max="${max + 1}"`);
      return `<col ${newAttrs}/>`;
    }
  );
  
  // 새 col B 삽입 (<cols> 바로 안쪽에)
  const newColB = `<col width="${newColBWidth}" customWidth="1" min="2" max="2"/>`;
  if (result.includes('<cols>')) {
    result = result.replace(/<cols>/, `<cols>${newColB}`);
  }
  // <cols>가 아예 없는 시트도 있음 - 그 경우는 패스 (스타일 없이 진행)
  
  return result;
}

// 원본 sheet XML에서 컬럼 A의 행별 스타일 ID 추출 (새 B 셀 추가 시 fallback)
function extractColAStyles(sheetXml) {
  const styles = new Map(); // rowNum(string) → styleId(string)
  const re = /<c\s+([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g;
  let m;
  while ((m = re.exec(sheetXml)) !== null) {
    const attrs = m[1];
    const refMatch = attrs.match(/\br="A(\d+)"/);
    if (!refMatch) continue;
    const styleMatch = attrs.match(/\bs="(\d+)"/);
    if (styleMatch) styles.set(refMatch[1], styleMatch[1]);
  }
  return styles;
}

// 시트의 <dimension ref="A1:J24"/>도 새 컬럼 추가에 맞춰 +1
function expandDimension(sheetXml) {
  return sheetXml.replace(
    /<dimension\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/,
    (m, startCol, startRow, endCol, endRow) => {
      const newEndCol = encodeColLetter(decodeColLetter(endCol) + 1);
      return `<dimension ref="${startCol}${startRow}:${newEndCol}${endRow}"/>`;
    }
  );
}

// 번역된 헤더 라벨
const HEADER_LABELS = {
  category: { en: 'Category', ja: 'カテゴリー', 'zh-TW': '類別', zh: '类别', th: 'หมวดหมู่', vi: 'Hạng mục' },
  eventName: { en: 'Event Name', ja: 'メニュー名', 'zh-TW': '項目名稱', zh: '项目名称', th: 'ชื่อรายการ', vi: 'Tên dịch vụ' },
  price: { en: 'Price', ja: '価格', 'zh-TW': '價格', zh: '价格', th: 'ราคา', vi: 'Giá' },
  description: { en: 'Description', ja: '説明', 'zh-TW': '說明', zh: '说明', th: 'รายละเอียด', vi: 'Mô tả' },
};

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXmlAttr(s) {
  return String(s)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

// sheet XML에서 지정된 셀들의 값만 inline string으로 치환.
// fallbackStyles: 새로 추가하는 셀의 경우 적용할 스타일 ID (cellRef → styleId)
function modifyCellsInSheetXml(sheetXml, modifications, fallbackStyles = null) {
  const processed = new Set();
  
  // Phase 1: 기존 셀 치환
  let result = sheetXml.replace(
    /<c\s+([^>]*?)\s*(\/>|>[\s\S]*?<\/c>)/g,
    (match, attrsStr) => {
      const refMatch = attrsStr.match(/\br="([^"]+)"/);
      if (!refMatch) return match;
      const ref = refMatch[1];
      if (!modifications.has(ref)) return match;
      
      const value = modifications.get(ref);
      processed.add(ref);
      
      const styleMatch = attrsStr.match(/\bs="([^"]*)"/);
      const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';
      
      return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
    }
  );
  
  // Phase 2: 못 찾은 셀들 → 행별로 그룹화 후 column 순 정렬 삽입
  // 중요: OOXML은 row 내 cell들이 column 순서로 정렬되어 있어야 Excel이 제대로 렌더링.
  // 그냥 끝에 붙이면 SheetJS는 읽지만 Excel은 셀을 표시 안 함.
  const newCellsByRow = new Map(); // rowNum(string) → [{col, xml}]
  for (const [ref, value] of modifications) {
    if (processed.has(ref)) continue;
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    const rowNum = m[2];
    const col = decodeColLetter(m[1]);
    const styleId = fallbackStyles?.get(ref);
    const styleAttr = styleId ? ` s="${styleId}"` : '';
    const cellXml = `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
    if (!newCellsByRow.has(rowNum)) newCellsByRow.set(rowNum, []);
    newCellsByRow.get(rowNum).push({ col, xml: cellXml });
  }
  
  // 각 행에 대해: 기존 셀 + 새 셀을 모두 column 순서로 재정렬
  for (const [rowNum, newCells] of newCellsByRow) {
    const rowRe = new RegExp(`(<row\\s+r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
    const rowMatch = result.match(rowRe);
    if (!rowMatch) continue;
    
    const rowOpen = rowMatch[1];
    const rowContent = rowMatch[2];
    const rowClose = rowMatch[3];
    
    // 행 내 기존 셀들을 추출 (col 인덱스와 함께)
    const existingCells = [];
    const cellRe = /<c\s+[^>]*?\s*(?:\/>|>[\s\S]*?<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rowContent)) !== null) {
      const refMatch = cm[0].match(/\br="([A-Z]+)\d+"/);
      if (!refMatch) continue;
      existingCells.push({ col: decodeColLetter(refMatch[1]), xml: cm[0] });
    }
    
    // 기존 + 새 셀 합쳐서 column 순 정렬
    const allCells = [...existingCells, ...newCells];
    allCells.sort((a, b) => a.col - b.col);
    const newRowContent = allCells.map(c => c.xml).join('');
    
    result = result.replace(rowOpen + rowContent + rowClose, rowOpen + newRowContent + rowClose);
  }
  
  return result;
}

// workbook.xml + workbook.xml.rels에서 시트 이름으로 worksheet XML 경로 찾기
async function findSheetXmlPath(filesMap, sheetName) {
  const decoder = new TextDecoder('utf-8');
  
  const wbFile = filesMap.get('xl/workbook.xml');
  if (!wbFile) throw new Error('xl/workbook.xml not found');
  const workbookXml = decoder.decode(await readXlsxFile(wbFile));
  
  const relsFile = filesMap.get('xl/_rels/workbook.xml.rels');
  if (!relsFile) throw new Error('xl/_rels/workbook.xml.rels not found');
  const relsXml = decoder.decode(await readXlsxFile(relsFile));
  
  // <sheet ... name="..." r:id="..."/> 찾기 - 속성 순서 무관
  const sheetRe = /<sheet\s+([^>]*?)\/?>/g;
  let rid = null;
  let m;
  while ((m = sheetRe.exec(workbookXml)) !== null) {
    const attrs = m[1];
    const nameMatch = attrs.match(/\bname="([^"]+)"/);
    const ridMatch = attrs.match(/\br:id="([^"]+)"/i);
    if (nameMatch && ridMatch && decodeXmlAttr(nameMatch[1]) === sheetName) {
      rid = ridMatch[1];
      break;
    }
  }
  if (!rid) throw new Error('Sheet not found: ' + sheetName);
  
  // <Relationship ... Id="..." ... Target="..."/> - 속성 순서 무관 (실제로는 Type/Target/Id 순)
  const relRe = /<Relationship\s+([^>]*?)\/?>/g;
  while ((m = relRe.exec(relsXml)) !== null) {
    const attrs = m[1];
    const idMatch = attrs.match(/\bId="([^"]+)"/);
    const targetMatch = attrs.match(/\bTarget="([^"]+)"/);
    if (idMatch && targetMatch && idMatch[1] === rid) {
      let target = targetMatch[1];
      if (target.startsWith('/')) target = target.slice(1);
      else if (!target.startsWith('xl/')) target = 'xl/' + target;
      return target;
    }
  }
  throw new Error('Relationship not found for ' + rid);
}

// 원본 디자인 그대로 보존하면서 번역된 값만 적용한 XLSX bytes 생성.
// 카테고리 열은 한국어(A) + 번역(B) 2열로 분리.
async function buildOutputBytes(sourceArrayBuffer, sourceParsed, translations, lang, db) {
  const files = await parseXlsxZip(sourceArrayBuffer);
  const sheetPath = await findSheetXmlPath(files, sourceParsed.sheetName);
  
  const decoder = new TextDecoder('utf-8');
  let sheetXml = decoder.decode(await readXlsxFile(files.get(sheetPath)));
  
  // 시프트 전에 컬럼 A의 행별 스타일 추출 (새 B 셀에 적용)
  const colAStyles = extractColAStyles(sheetXml);
  
  // === Phase 1: 데이터 영역(헤더 행 포함) B+ 컬럼 → C+ 시프트 ===
  const headerRow1 = sourceParsed.headerRow + 1; // 0-indexed → 1-indexed
  sheetXml = shiftDataColumnsRight(sheetXml, headerRow1);
  sheetXml = shiftColsAndInsertNewB(sheetXml);
  sheetXml = expandDimension(sheetXml);
  
  // === Phase 2: 셀 값 채우기 ===
  // 시프트 후 새 컬럼 배치:
  //   A = 한국어 카테고리 (원본 유지)
  //   B = 번역된 카테고리 (NEW)
  //   C = 번역된 이벤트명 (시프트된 B 자리)
  //   D = 가격 (시프트된 C 자리)
  //   E = 번역된 설명 (시프트된 D 자리)
  const catMap = db.cat || {};
  const modifications = new Map();
  const fallbackStyles = new Map(); // 새로 추가될 셀의 스타일
  const headerRow0 = sourceParsed.headerRow;
  
  // 헤더 행 번역
  const catHeader = HEADER_LABELS.category[lang] || 'Category';
  const evtHeader = HEADER_LABELS.eventName[lang] || 'Event Name';
  const priceHeader = HEADER_LABELS.price[lang] || 'Price';
  const descHeader = HEADER_LABELS.description[lang] || 'Description';
  
  modifications.set(encodeCellRef(headerRow0, 1), catHeader);        // B6: 새 카테고리 헤더
  modifications.set(encodeCellRef(headerRow0, 2), evtHeader);        // C6: 이벤트명 (was B)
  modifications.set(encodeCellRef(headerRow0, 3), `${priceHeader} (₩)`); // D6: 가격
  modifications.set(encodeCellRef(headerRow0, 4), descHeader);       // E6: 설명
  
  // 헤더 행의 새 B 셀 스타일 = 원본 A 헤더 셀 스타일
  const headerStyleId = colAStyles.get(String(headerRow1));
  if (headerStyleId) {
    fallbackStyles.set(encodeCellRef(headerRow0, 1), headerStyleId);
  }
  
  // 데이터 행
  for (const event of sourceParsed.events) {
    const trans = translations[event.id];
    if (!trans) continue;
    const rowIdx = event.rowIdx;
    
    // 카테고리: A는 원본 한국어 그대로 (수정 안 함), B에 번역 추가
    if (event.categoryShownInThisRow) {
      const catTrans = catMap[event.category];
      if (catTrans && catTrans[lang]) {
        const refB = encodeCellRef(rowIdx, 1);
        modifications.set(refB, catTrans[lang]);
        // A의 스타일 복사
        const aStyle = colAStyles.get(String(rowIdx + 1));
        if (aStyle) fallbackStyles.set(refB, aStyle);
      }
    }
    
    // C: 번역된 이벤트명
    if (trans.event) {
      modifications.set(encodeCellRef(rowIdx, 2), trans.event);
    }
    
    // D: 가격
    if (event.price !== null && event.price !== undefined && event.price !== '') {
      const priceFormatted = formatPrice(event.price);
      if (priceFormatted) {
        modifications.set(encodeCellRef(rowIdx, 3), priceFormatted);
      }
    }
    
    // E: 번역된 설명
    if (trans.desc) {
      modifications.set(encodeCellRef(rowIdx, 4), trans.desc);
    }
  }
  
  // === Phase 3: 컬럼 B 시각 통일성 보장 ===
  // 시프트로 인해 새 B 컬럼은 셀이 없어 흰 배경으로 표시됨.
  // 헤더 행 이하의 모든 styled A 셀에 대해, B에 modification이 없으면
  // 빈 B 셀에 A의 스타일을 복사하여 디자인 연속성 유지.
  for (const [rowNumStr, styleId] of colAStyles) {
    const rowIdx0 = parseInt(rowNumStr, 10) - 1;
    if (rowIdx0 < sourceParsed.headerRow) continue; // 헤더 이전 (제목/안내 영역) 스킵
    const refB = encodeCellRef(rowIdx0, 1);
    if (modifications.has(refB)) continue; // 이미 값 있는 셀(카테고리 번역 등)은 그대로
    modifications.set(refB, '');
    fallbackStyles.set(refB, styleId);
  }
  
  sheetXml = modifyCellsInSheetXml(sheetXml, modifications, fallbackStyles);
  
  // 나머지 ZIP 재구성
  const rawFiles = new Map();
  for (const [name, entry] of files) {
    if (name === sheetPath) {
      rawFiles.set(name, new TextEncoder().encode(sheetXml));
    } else {
      rawFiles.set(name, await readXlsxFile(entry));
    }
  }
  
  return await buildXlsxZip(rawFiles);
}

// XLSX write 옵션 - cellStyles로 디자인(색상, 폰트, 정렬 등) 보존
const XLSX_WRITE_OPTS = { type: 'array', bookType: 'xlsx', cellStyles: true };

function workbookToBuffer(wb) {
  return new Uint8Array(XLSX.write(wb, XLSX_WRITE_OPTS));
}

function downloadWorkbook(wb, filename) {
  const buf = workbookToBuffer(wb);
  triggerDownload(buf, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

function triggerDownload(bytes, filename, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// === 인라인 ZIP 라이터 (STORED 메소드, 무압축) ===
// XLSX 자체가 이미 ZIP 압축되어 있어서 외부 ZIP은 STORED로도 충분
const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(data) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    c = CRC32_TABLE[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function createZipBlob(files) {
  // files: [{ name: string, data: Uint8Array }]
  const encoder = new TextEncoder();
  const parts = [];
  const centralParts = [];
  let offset = 0;
  
  for (const file of files) {
    const nameBytes = encoder.encode(file.name);
    const data = file.data;
    const crc = crc32(data);
    const size = data.length;
    
    // Local file header (30 bytes)
    const lfh = new Uint8Array(30);
    const lfhView = new DataView(lfh.buffer);
    lfhView.setUint32(0, 0x04034b50, true);
    lfhView.setUint16(4, 20, true);
    lfhView.setUint16(6, 0x0800, true); // UTF-8 filename
    lfhView.setUint16(8, 0, true);       // stored
    lfhView.setUint16(10, 0, true);
    lfhView.setUint16(12, 0, true);
    lfhView.setUint32(14, crc, true);
    lfhView.setUint32(18, size, true);
    lfhView.setUint32(22, size, true);
    lfhView.setUint16(26, nameBytes.length, true);
    lfhView.setUint16(28, 0, true);
    
    parts.push(lfh);
    parts.push(nameBytes);
    parts.push(data);
    
    // Central directory entry (46 bytes)
    const cde = new Uint8Array(46);
    const cdeView = new DataView(cde.buffer);
    cdeView.setUint32(0, 0x02014b50, true);
    cdeView.setUint16(4, 20, true);
    cdeView.setUint16(6, 20, true);
    cdeView.setUint16(8, 0x0800, true);
    cdeView.setUint16(10, 0, true);
    cdeView.setUint16(12, 0, true);
    cdeView.setUint16(14, 0, true);
    cdeView.setUint32(16, crc, true);
    cdeView.setUint32(20, size, true);
    cdeView.setUint32(24, size, true);
    cdeView.setUint16(28, nameBytes.length, true);
    cdeView.setUint16(30, 0, true);
    cdeView.setUint16(32, 0, true);
    cdeView.setUint16(34, 0, true);
    cdeView.setUint16(36, 0, true);
    cdeView.setUint32(38, 0, true);
    cdeView.setUint32(42, offset, true);
    
    centralParts.push(cde);
    centralParts.push(nameBytes);
    offset += 30 + nameBytes.length + size;
  }
  
  const centralStart = offset;
  let centralSize = 0;
  for (const p of centralParts) {
    parts.push(p);
    centralSize += p.length;
  }
  
  // End of central directory (22 bytes)
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, files.length, true);
  eocdView.setUint16(10, files.length, true);
  eocdView.setUint32(12, centralSize, true);
  eocdView.setUint32(16, centralStart, true);
  eocdView.setUint16(20, 0, true);
  parts.push(eocd);
  
  // Concatenate
  let total = 0;
  for (const p of parts) total += p.length;
  const result = new Uint8Array(total);
  let pos = 0;
  for (const p of parts) {
    result.set(p, pos);
    pos += p.length;
  }
  return result;
}

function getOutputFilename(srcFilename, lang) {
  const langSuffix = {
    'en': '_영어',
    'ja': '_일본어',
    'zh-TW': '_대만어',
    'zh': '_중국어',
    'th': '_태국어',
    'vi': '_베트남어',
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
    en: true, ja: false, 'zh-TW': false, zh: false, th: false, vi: true,
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
      
      // 원본 파일의 언어 선택 자동 적용 (감지된 것만 = true, 나머지 = false)
      // 감지 실패 시 (langFlags 전부 false) → 모두 비선택 상태로 표시하여
      // 사용자가 수동으로 골라야 함을 명확히
      const newSel = {};
      LANGUAGES.forEach(l => { newSel[l.code] = !!data.langFlags[l.code]; });
      setSelectedLangs(newSel);
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
      
      // 3. 출력 XLSX bytes 생성 (디자인 100% 보존 - ZIP+XML 방식)
      setProgress({ phase: 'building', total: targetLangs.length, done: 0 });
      const newOutputs = {};
      for (let i = 0; i < targetLangs.length; i++) {
        const lang = targetLangs[i];
        const bytes = await buildOutputBytes(parsed.arrayBuffer, parsed, allTranslations[lang], lang, DATA);
        const filename = getOutputFilename(parsed.filename, lang);
        newOutputs[lang] = { bytes, filename };
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
    if (out) triggerDownload(out.bytes, out.filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  };
  
  const handleDownloadAll = () => {
    const entries = Object.entries(outputs);
    if (entries.length === 0) return;
    // 이미 생성된 bytes를 ZIP으로 묶기
    const files = entries.map(([lang, out]) => ({
      name: out.filename,
      data: out.bytes,
    }));
    const zipBytes = createZipBlob(files);
    // 원본 파일명에서 확장자 + 언어 접미사 떼고 _번역_전체.zip
    let baseName = (parsed?.filename || '미미썸_이벤트').replace(/\.[^.]+$/, '');
    const langSuffixes = ['_영어', '_일본어', '_대만어', '_중국어', '_태국어', '_베트남어', '_한국어'];
    for (const suffix of langSuffixes) {
      if (baseName.endsWith(suffix)) {
        baseName = baseName.slice(0, -suffix.length);
        break;
      }
    }
    triggerDownload(zipBytes, `${baseName}_번역_전체.zip`, 'application/zip');
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

          {/* 3. Language selection - 자동 감지됨 */}
          <div className="p-5 rounded-lg" style={{ backgroundColor: '#FFFEFC', border: '1px solid #E8DFD0' }}>
            <div className="flex items-center justify-between mb-3">
              <h2 className="font-display text-lg font-medium" style={{ color: '#3D2817' }}>
                <span style={{ color: '#B8804E' }}>→</span> 번역할 언어
              </h2>
              {(() => {
                const detected = Object.values(parsed.langFlags).filter(Boolean).length;
                if (detected > 0) {
                  return (
                    <span className="text-xs px-2.5 py-1 rounded-full font-medium" style={{ backgroundColor: '#E8F0E0', color: '#4A6B3A' }}>
                      ✓ 원본 시트에서 {detected}개 언어 자동 감지
                    </span>
                  );
                }
                return (
                  <span className="text-xs px-2.5 py-1 rounded-full font-medium" style={{ backgroundColor: '#FBEBD7', color: '#8B5A2B' }}>
                    ⚠ 원본 시트에서 선택된 언어 없음 — 아래에서 직접 선택해주세요
                  </span>
                );
              })()}
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
            <p className="text-[11px]" style={{ color: '#86715A' }}>
              💡 원본 엑셀 상단에서 ☑로 체크한 언어가 자동 선택됩니다. 필요하면 위 타일을 클릭해서 변경할 수 있어요.
            </p>
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
                  전체 ZIP 다운로드
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
    en: true, ja: true, 'zh-TW': true, zh: true, th: true, vi: true,
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
            const fontClass = langCode === 'ja' ? 'font-ja' : langCode === 'zh' ? 'font-zh' : langCode === 'zh-TW' ? 'font-zh-tw' : langCode === 'th' ? 'font-th' : '';
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

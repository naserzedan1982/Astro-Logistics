import pandas as pd
from openpyxl.styles import PatternFill, Font
import os

# 1. إنشاء بيانات المحاكاة (Mock Data) كبديل للمخططات الهندسية
data = {
    'Item_ID': [
        'N2-TANK-MAIN-01', 'O2-SPOOL-001', 'HABITAT-MOD-A', 'COOLING-PIPE-X', 'REACTOR-CORE-01',
        'SOLAR-ARRAY-01', 'SOLAR-ARRAY-02', 'LUNAR-ROVER-MK2', 'AIRLOCK-SYS-A', 'WATER-REC-UNIT',
        'COMM-RELAY-TWR', 'O2-TANK-AUX-01', 'N2-SPOOL-044', 'THERMAL-RAD-01', 'STRUCT-BEAM-H',
        'HABITAT-MOD-B', 'H2-FUEL-CELL', 'COOLING-PUMP', 'PIPING-MANIFOLD', 'FOUNDATION-PAD'
    ],
    'System_Category': [
        'Nitrogen Storage', 'Life Support Piping', 'Structural', 'Cooling System', 'Power Generation',
        'Power Generation', 'Power Generation', 'Transport', 'Life Support', 'Life Support',
        'Communications', 'Oxygen Storage', 'Life Support Piping', 'Cooling System', 'Structural',
        'Structural', 'Power Generation', 'Cooling System', 'Piping System', 'Structural'
    ],
    'Isometric_Ref': [
        'DWG-ISO-405', 'DWG-ISO-102', 'DWG-STR-01', 'DWG-ISO-205', 'DWG-EQP-900',
        'DWG-PWR-101', 'DWG-PWR-102', 'DWG-TRN-001', 'DWG-LFS-303', 'DWG-LFS-401',
        'DWG-COM-505', 'DWG-ISO-406', 'DWG-ISO-103', 'DWG-ISO-206', 'DWG-STR-02',
        'DWG-STR-03', 'DWG-PWR-202', 'DWG-ISO-207', 'DWG-ISO-500', 'DWG-STR-00'
    ],
    'Earth_Weight_Ton': [
        170.0, 5.5, 45.0, 3.2, 120.0,
        12.5, 12.5, 8.0, 18.0, 9.2,
        4.5, 145.0, 2.1, 11.0, 22.0,
        45.0, 35.0, 6.4, 4.8, 85.0
    ],
    'Dimensions_Meters': [
        '38x8x8', '6x0.5x0.5', '10x5x5', '12x0.3x0.3', '15x10x10',
        '20x4x1', '20x4x1', '4x3x2', '5x4x4', '3x3x3',
        '15x2x2', '32x7x7', '4x0.5x0.5', '8x8x1', '10x2x2',
        '10x5x5', '8x4x4', '2x2x2', '5x3x3', '20x20x2'
    ],
    'Lifting_Rigging_Type': [
        'Heavy Crawler Crane', 'Manual/Light Crane', 'Medium Crane', 'Manual', 'Heavy Crawler Crane',
        'Medium Crane', 'Medium Crane', 'Self-Propelled', 'Medium Crane', 'Light Crane',
        'Light Crane', 'Heavy Crawler Crane', 'Manual', 'Medium Crane', 'Heavy Crawler Crane',
        'Medium Crane', 'Heavy Crawler Crane', 'Light Crane', 'Light Crane', 'Heavy Crawler Crane'
    ],
    'Installation_Phase': [
        'Phase 1: Mobilization', 'Phase 2: Piping Setup', 'Phase 1: Mobilization', 'Phase 2: Piping Setup', 'Phase 1: Mobilization',
        'Phase 1: Mobilization', 'Phase 1: Mobilization', 'Phase 1: Mobilization', 'Phase 2: Assembly', 'Phase 2: Assembly',
        'Phase 1: Mobilization', 'Phase 1: Mobilization', 'Phase 2: Piping Setup', 'Phase 2: Assembly', 'Phase 1: Mobilization',
        'Phase 1: Mobilization', 'Phase 1: Mobilization', 'Phase 2: Assembly', 'Phase 2: Piping Setup', 'Phase 1: Mobilization'
    ],
    'Status': [
        'In-Orbit Transit', 'Manufactured', 'Landed - Ready', 'In-Orbit Transit', 'Landed - Pending Lift',
        'Landed - Ready', 'Landed - Ready', 'Operational', 'In-Orbit Transit', 'Manufactured',
        'Landed - Ready', 'Manufactured', 'Manufactured', 'In-Orbit Transit', 'Landed - Pending Lift',
        'Manufactured', 'In-Orbit Transit', 'Landed - Ready', 'Manufactured', 'Landed - Ready'
    ]
}

# تحويل البيانات إلى DataFrame
df = pd.DataFrame(data)

# 2. محرك المنطق الهندسي (Engineering Logic Engine)
# حساب الوزن على القمر (سدس وزن الأرض)
df['Lunar_Weight_Ton'] = (df['Earth_Weight_Ton'] / 6).round(2)

# فرز البيانات: أعمال الرفع الثقيل (Heavy Lifting) وأعمال الأنابيب (Piping)
heavy_lifting_df = df[df['Earth_Weight_Ton'] > 20].sort_values(by='Earth_Weight_Ton', ascending=False)
piping_spools_df = df[df['System_Category'].str.contains('Piping|Cooling')]

# إنشاء بيانات لوحة القيادة (Dashboard)
dashboard_data = {
    'Metric': ['Total Earth Weight (Tons)', 'Total Lunar Weight (Tons)', 'Heavy Lifts Scheduled', 'Piping Spools Pending'],
    'Value': [df['Earth_Weight_Ton'].sum(), df['Lunar_Weight_Ton'].sum(), len(heavy_lifting_df), len(piping_spools_df)],
    'Status': ['Safe for Starship Launch', 'Optimal for Lunar Cranes', 'Requires Rigging Plan', 'Ready for Welding']
}
dashboard_df = pd.DataFrame(dashboard_data)

# 3. تصدير وتنسيق ملف الـ Excel باستخدام openpyxl
excel_filename = "Lunar_Site_Mobilization_Tracker.xlsx"

with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
    # كتابة الصفحات (Sheets)
    dashboard_df.to_excel(writer, sheet_name='Project_Dashboard', index=False)
    heavy_lifting_df.to_excel(writer, sheet_name='Heavy_Lifting_Schedule', index=False)
    piping_spools_df.to_excel(writer, sheet_name='Piping_and_Spools', index=False)
    df.to_excel(writer, sheet_name='Raw_Isometric_Data', index=False)

    # الوصول إلى ملف العمل لتنسيقه
    workbook = writer.book
    
    # تنسيق رأس الجدول بالخط العريض
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # أزرق ناسا
    warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # أحمر للتحذير
    
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        
        # تنسيق صف الرأس
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            
        # توسيع الأعمدة لتناسب النص
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

    # إضافة تنبيه (Conditional Formatting مبسط) للمعدات الثقيلة جداً في صفحة الرفع
    heavy_sheet = workbook['Heavy_Lifting_Schedule']
    # نفترض أن عمود الوزن القمري هو H
    for row in range(2, heavy_sheet.max_row + 1):
        lunar_weight_cell = heavy_sheet[f'I{row}'] # العمود التاسع
        if isinstance(lunar_weight_cell.value, (int, float)) and lunar_weight_cell.value > 25:
            lunar_weight_cell.fill = warning_fill # تلوين بالأحمر إذا تجاوز الوزن القمري 25 طن

print(f"تم إنشاء الملف بنجاح: {os.path.abspath(excel_filename)}")
print("افتح الملف لتتفقد خريطة التعبئة وجدول الرفع الثقيل!")

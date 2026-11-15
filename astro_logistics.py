import pandas as pd
from openpyxl.styles import PatternFill, Font
import os

# 1. إنشاء بيانات المحاكاة (Mock Data) كبديل للمخططات الهندسية
data = {
    'Item_ID': ['N2-TANK-MAIN-01', 'O2-SPOOL-001', 'HABITAT-MOD-A', 'COOLING-PIPE-X', 'REACTOR-CORE-01'],
    'System_Category': ['Nitrogen Storage', 'Life Support Piping', 'Structural', 'Cooling System', 'Power Generation'],
    'Isometric_Ref': ['DWG-LUNAR-ISO-405', 'DWG-LUNAR-ISO-102', 'DWG-LUNAR-STR-01', 'DWG-LUNAR-ISO-205', 'DWG-LUNAR-EQP-900'],
    'Earth_Weight_Ton': [170.0, 5.5, 45.0, 3.2, 120.0], # لاحظ وزن خزان النيتروجين 170 طن
    'Dimensions_Meters': ['38 x 8 x 8', '6 x 0.5 x 0.5', '10 x 5 x 5', '12 x 0.3 x 0.3', '15 x 10 x 10'],
    'Lifting_Rigging_Type': ['Heavy Crawler Crane', 'Manual/Light Crane', 'Medium Crane', 'Manual', 'Heavy Crawler Crane'],
    'Installation_Phase': ['Phase 1: Site Mobilization', 'Phase 2: Piping Setup', 'Phase 1: Site Mobilization', 'Phase 2: Piping Setup', 'Phase 1: Site Mobilization'],
    'Status': ['In-Orbit Transit', 'Manufactured', 'Landed - Ready', 'In-Orbit Transit', 'Landed - Pending Lift']
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
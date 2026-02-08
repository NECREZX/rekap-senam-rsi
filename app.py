from flask import Flask, render_template, jsonify, request, send_file, send_from_directory
import pandas as pd
import os
import json
import tempfile
import shutil
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import io
import numpy as np
from collections import OrderedDict
import math

app = Flask(__name__, static_folder='static', template_folder='templates')

# Konfigurasi
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
DATA_FILE = os.path.join(BASE_DIR, 'data/senam_data.json')
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
TEMP_DATA_FILE = os.path.join(BASE_DIR, 'data/current_upload.json')

# Buat folder
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('static/css', exist_ok=True)
os.makedirs('static/js', exist_ok=True)
os.makedirs('data', exist_ok=True)
os.makedirs('templates', exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def normalize_col(col):
    if pd.isna(col):
        return ""
    return str(col).strip().replace(" ", "").upper()

def parse_month_year(col_name):
    """Parse bulan dan tahun dari nama kolom"""
    col_str = str(col_name)
    import re
    
    patterns = [
        r'(202[2-9]|203[0-2])[-/](\d{1,2})',
        r'(202[2-9]|203[0-2])(\d{2})',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, col_str)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            if 1 <= int(month) <= 12:
                return year, month
    
    return None, None

def excel_to_json(file_path):
    """Konversi Excel/CSV ke format JSON"""
    try:
        if file_path.endswith('.csv'):
            encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
            df = None
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, header=4, encoding=encoding)
                    break
                except:
                    continue
            if df is None:
                raise ValueError("Gagal membaca file CSV")
        else:
            df = pd.read_excel(file_path, header=4)
        
        if df.empty or len(df.columns) < 3:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, header=None)
            else:
                df = pd.read_excel(file_path, header=None)
            
            header_row = None
            for idx in range(min(10, len(df))):
                row_values = df.iloc[idx].astype(str).str.upper().tolist()
                if 'NAMA' in row_values:
                    header_row = idx
                    break
            
            if header_row is not None:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, header=header_row, encoding='utf-8')
                else:
                    df = pd.read_excel(file_path, header=header_row)
        
        df.columns = [normalize_col(c) for c in df.columns]
        
        result = []
        years = [str(year) for year in range(2022, 2033)]
        
        monthly_cols = {}
        for year in years:
            monthly_cols[year] = {}
        
        month_names = {
            '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
            '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
            '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
        }
        
        for col in df.columns:
            col_str = str(col)
            year, month = parse_month_year(col_str)
            if year and month:
                if year in monthly_cols:
                    monthly_cols[year][col] = month
        
        for idx, row in df.iterrows():
            nama = str(row.get("NAMA", "")).strip()
            if not nama or nama.lower() in ['', 'nan', 'none', 'null']:
                continue

            nik_value = str(row.get("NIK", "")).strip()
            unique_id = f"{nik_value}_{idx}" if nik_value else f"emp_{idx}"
            
            pegawai = {
                "id": unique_id,
                "original_index": idx,
                "nama": nama,
                "nik": str(row.get("NIK", "")).strip(),
                "jk": str(row.get("JK", "")).strip(),
                "status": str(row.get("STATUSPEGAWAI", "")).strip(),
                "kelompok": str(row.get("KELOMPOKNAKES", "")).strip(),
                "jabatan": str(row.get("NAMAJABATAN", "")).strip(),
                "struktur": str(row.get("STRUKTURLINI", "")).strip(),
                "tempat": str(row.get("TEMPATTUGAS", "")).strip(),
                "keterangan": str(row.get("KETERANGANUNTUKPEMANGGILAN", "") if "KETERANGANUNTUKPEMANGGILAN" in df.columns else "").strip(),
                "bulanan": {},
                "tahunan": {},
                "total_all": 0,
                "shift_status": "non_shift"
            }
            
            for year in years:
                pegawai["bulanan"][year] = OrderedDict()
                for month_num in range(1, 13):
                    month_key = f"{year}-{str(month_num).zfill(2)}"
                    pegawai["bulanan"][year][month_key] = {
                        "nama": month_names.get(str(month_num).zfill(2), f"Bulan {month_num}"),
                        "value": 0,
                        "status": "Tidak Hadir"
                    }
            
            total_all = 0
            yearly_totals = {year: 0 for year in years}
            
            for year in years:
                if year in monthly_cols:
                    for col_name, month_num in monthly_cols[year].items():
                        if col_name in df.columns:
                            month_key = f"{year}-{month_num}"
                            val = row[col_name]
                            
                            value = 0
                            status = "Tidak Hadir"
                            
                            if pd.isna(val):
                                value = 0
                                status = "Tidak Ada Data"
                            elif isinstance(val, (int, float)):
                                value = int(val)
                                status = "Hadir" if value > 0 else "Tidak Hadir"
                            elif isinstance(val, str):
                                val_lower = val.lower().strip()
                                if val_lower in ['senam', 'hadir', 'ya', 'y', '1', 'v', '✓', '✔', 'hadir senam']:
                                    value = 1
                                    status = "Hadir"
                                elif val_lower in ['tidak', 'tidak hadir', 'no', 'n', '0', 'x', '✗', '❌', '-', '']:
                                    value = 0
                                    status = "Tidak Hadir"
                                elif 'hamil' in val_lower:
                                    value = 0
                                    status = "Sedang Hamil"
                                elif 'cuti' in val_lower:
                                    value = 0
                                    status = "Sedang Cuti"
                                elif 'pelatihan' in val_lower:
                                    value = 0
                                    status = "Sedang Pelatihan"
                                else:
                                    try:
                                        num_val = float(val)
                                        value = int(num_val)
                                        status = "Hadir" if value > 0 else "Tidak Hadir"
                                    except:
                                        value = 0
                                        status = val
                            else:
                                value = 0
                                status = "Tidak Hadir"
                            
                            yearly_totals[year] += value
                            
                            pegawai["bulanan"][year][month_key] = {
                                "nama": month_names.get(month_num, f"Bulan {month_num}"),
                                "value": value,
                                "status": status
                            }
                
                total_cols = [c for c in df.columns if f"JUMLAH{year}" in c or f"TOTAL{year}" in c or f"{year}TOTAL" in c]
                if total_cols:
                    for col in total_cols:
                        if col in df.columns:
                            val = row[col]
                            if not pd.isna(val):
                                try:
                                    if isinstance(val, (int, float)):
                                        yearly_totals[year] = int(val)
                                    elif isinstance(val, str):
                                        yearly_totals[year] = int(float(val.strip()))
                                except:
                                    pass
            
            for year in years:
                pegawai["tahunan"][year] = yearly_totals[year]
                total_all += yearly_totals[year]
            
            pegawai["total_all"] = total_all
            result.append(pegawai)
        
        return result
    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        traceback.print_exc()
        return None

def save_temp_data(data):
    try:
        with open(TEMP_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                'data': data,
                'timestamp': datetime.now().isoformat(),
                'count': len(data)
            }, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error saving temp data: {e}")
        return False

def load_temp_data():
    try:
        if os.path.exists(TEMP_DATA_FILE):
            with open(TEMP_DATA_FILE, 'r', encoding='utf-8') as f:
                content = json.load(f)
                timestamp = datetime.fromisoformat(content['timestamp'])
                if datetime.now() - timestamp < timedelta(hours=24):
                    return content['data']
        return []
    except Exception as e:
        print(f"Error loading temp data: {e}")
        return []

def clear_temp_data():
    try:
        if os.path.exists(TEMP_DATA_FILE):
            os.remove(TEMP_DATA_FILE)
        return True
    except:
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def api_data():
    data = load_temp_data()
    return jsonify(data)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'Tidak ada file yang diupload'})
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'message': 'Nama file kosong'})
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'message': 'Format file tidak didukung'})
        
        file.seek(0, 2)
        file_size = file.tell()
        file.seek(0)
        if file_size > 10 * 1024 * 1024:
            return jsonify({'success': False, 'message': 'Ukuran file terlalu besar'})
        
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(temp_path)
        
        clear_temp_data()
        
        data = excel_to_json(temp_path)
        
        if data is None:
            shutil.rmtree(temp_dir)
            return jsonify({'success': False, 'message': 'Gagal memproses file'})
        
        if not data:
            shutil.rmtree(temp_dir)
            return jsonify({'success': False, 'message': 'Tidak ada data'})
        
        if save_temp_data(data):
            shutil.rmtree(temp_dir)
            return jsonify({
                'success': True,
                'message': f'Data berhasil diupload! {len(data)} pegawai diproses.',
                'count': len(data),
                'years': list(data[0]['tahunan'].keys()) if data else []
            })
        else:
            shutil.rmtree(temp_dir)
            return jsonify({'success': False, 'message': 'Gagal menyimpan data'})
            
    except Exception as e:
        print(f"Upload error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})

@app.route('/api/export-pdf', methods=['POST'])
def export_pdf():
    """Export PDF 1 HALAMAN A4 PORTRAIT - Format Lengkap"""
    try:
        data = request.json
        employee_data = data.get('employee_data')
        date_range = data.get('date_range', {})
        shift_status = data.get('shift_status', 'non_shift')
        bulanan_data = data.get('bulanan_data', {})
        selected_year = data.get('selected_year', '2024')
        
        if not employee_data:
            return jsonify({'success': False, 'message': 'Data tidak ditemukan'})
        
        # Sort data bulanan
        sorted_months = sorted(bulanan_data.items(), key=lambda x: x[0])
        
        # Hitung analisis
        total_attendance = sum(month_data.get('value', 0) for _, month_data in sorted_months)
        months_count = len(sorted_months)
        
        # TARGET FIXED - Tidak berubah meskipun jumlah bulan berbeda
        # Shift: 20 kali (50% dari 40)
        # Non-Shift: 28 kali (70% dari 40)
        if shift_status == 'shift':
            target_attendance = 40  # FIXED
            percentage_target = 50  # Untuk tampilan di PDF
        else:
            target_attendance = 56  # FIXED
            percentage_target = 70  # Untuk tampilan di PDF
        
        status_tercapai = total_attendance >= target_attendance
        
        # CREATE PDF - A4 PORTRAIT
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,  # A4 Portrait
            rightMargin=1.2*cm,
            leftMargin=1.2*cm,
            topMargin=0.8*cm,
            bottomMargin=0.8*cm
        )
        
        story = []
        styles = getSampleStyleSheet()
        
        # CUSTOM STYLES - Disesuaikan untuk A4
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=11,  # Dikurangi dari 12
            alignment=TA_CENTER,
            spaceAfter=3,
            textColor=colors.HexColor('#1a237e'),
            fontName='Helvetica-Bold'
        )
        
        subtitle_style = ParagraphStyle(
            'SubtitleStyle',
            parent=styles['Heading2'],
            fontSize=9,  # Dikurangi dari 10
            alignment=TA_CENTER,
            spaceAfter=6,
            textColor=colors.HexColor('#283593'),
            fontName='Helvetica-Bold'
        )
        
        # HEADER
        story.append(Paragraph("RUMAH SAKIT ISLAM SITI KHADIJAH PALEMBANG", title_style))
        story.append(Paragraph("REKAP ABSENSI SENAM PEGAWAI", subtitle_style))
        story.append(Spacer(1, 0.1*cm))
        
        # Hitung tahun yang BENAR dari data bulanan
        if sorted_months:
            years = [month_key.split('-')[0] for month_key, _ in sorted_months]
            min_year = min(years)
            max_year = max(years)
            
            if min_year == max_year:
                period_text = min_year
            else:
                period_text = f"{min_year}-{max_year}"
        else:
            period_text = selected_year
        
        # INFO PEGAWAI - 2 KOLOM (Lebar disesuaikan A4)
        info_data = [
            ['NAMA', ':', employee_data.get('nama', '-'), 'JABATAN', ':', employee_data.get('jabatan', '-')],
            ['NIK', ':', employee_data.get('nik', '-'), 'TEMPAT TUGAS', ':', employee_data.get('tempat', '-')],
            ['STATUS SHIFT', ':', 'SHIFT' if shift_status == 'shift' else 'NON-SHIFT', 'TAHUN', ':', period_text]
        ]
        
        if date_range.get('start') and date_range.get('end'):
            info_data.append(['PERIODE', ':', f"{date_range['start']} s/d {date_range['end']}", '', '', ''])
        
        # Lebar total: 2.2+0.3+4.5+2.2+0.3+4.5 = 14cm (muat A4)
        info_table = Table(info_data, colWidths=[2.2*cm, 0.3*cm, 4.5*cm, 2.2*cm, 0.3*cm, 4.5*cm])
        info_table.setStyle(TableStyle([
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (3,0), (3,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 6.5),  # Font dikecilkan
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 2),
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 0.25*cm))
        
        # TABEL KEHADIRAN - TETAP 3 KOLOM dengan STATUS (Lebar disesuaikan)
        if len(sorted_months) > 0:
            rows_per_col = math.ceil(len(sorted_months) / 3)
            
            table_data = []
            for row_idx in range(rows_per_col):
                row = []
                for col_idx in range(3):
                    month_idx = col_idx * rows_per_col + row_idx
                    if month_idx < len(sorted_months):
                        month_key, month_data = sorted_months[month_idx]
                        # Singkat nama bulan: Jan 2024, Feb 2024, dll
                        month_name = month_data.get('nama', '')[:3] + ' ' + month_key.split('-')[0]
                        value = month_data.get('value', 0)
                        status = month_data.get('status', 'Tidak Hadir')
                        
                        row.extend([month_name, str(value), status])
                    else:
                        row.extend(['', '', ''])
                table_data.append(row)
            
            # Header
            header = ['BULAN', 'HADIR', 'KETERANGAN'] * 3
            table_data.insert(0, header)
            
            # Lebar disesuaikan A4: (2.5+0.8+2) * 3 = 15.9cm (muat!)
            col_widths = [2.5*cm, 0.8*cm, 2*cm] * 3
            
            attendance_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            attendance_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1976d2')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 6),  # Font header kecil
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('FONTSIZE', (0,1), (-1,-1), 5.5),  # Font isi kecil
                ('ALIGN', (0,1), (-1,-1), 'LEFT'),
                ('ALIGN', (1,1), (1,-1), 'CENTER'),
                ('ALIGN', (4,1), (4,-1), 'CENTER'),
                ('ALIGN', (7,1), (7,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')]),
                ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
                ('TOPPADDING', (0,0), (-1,-1), 1.5),
            ]))
            
            story.append(attendance_table)
            story.append(Spacer(1, 0.2*cm))
        
        # ANALISIS KEHADIRAN
        analysis_title = ParagraphStyle(
            'AnalysisTitle',
            parent=styles['Heading3'],
            fontSize=9,
            textColor=colors.HexColor('#1565c0'),
            spaceAfter=4,
            fontName='Helvetica-Bold'
        )
        
        story.append(Paragraph("", analysis_title))
        
        # Gunakan perhitungan target yang benar
        percentage_target = 50 if shift_status == 'shift' else 70
        
        analysis_data = [
            ['KETERANGAN', 'NILAI'],
            ['Total Kegiatan Senam (2 Tahun)', '80 kali'],
            [f'Target Kehadiran ({percentage_target}%)', f'{target_attendance} kali'],
            ['Total Kehadiran Aktual', f'{total_attendance} kali'],
            ['Persentase Kehadiran', f'{(total_attendance/target_attendance*100) if target_attendance > 0 else 0:.1f}%'],
            ['Selisih dari Target', f'{total_attendance - target_attendance:+d} kali']
        ]
        
        # Lebar: 9+5 = 14cm (muat A4)
        analysis_table = Table(analysis_data, colWidths=[9*cm, 5*cm])
        analysis_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e3f2fd')),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 7),  # Font kecil
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (1,0), (1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#fafafa')]),
        ]))
        
        story.append(analysis_table)
        story.append(Spacer(1, 0.2*cm))
        
        # KESIMPULAN
        if status_tercapai:
            conclusion_text = f"""
            <b>✓ KESIMPULAN: TELAH MENCUKUPI TARGET</b><br/>
            Pegawai <b>TELAH MEMENUHI</b> target kehadiran senam.<br/>
            Total kehadiran: <b>{total_attendance} kali</b> dari target <b>{target_attendance} kali</b><br/>
            Persentase pencapaian: <b>{(total_attendance/target_attendance*100) if target_attendance > 0 else 0:.1f}%</b> dari target {percentage_target}%
            """
            conclusion_color = colors.HexColor('#2e7d32')
            bg_color = colors.HexColor('#e8f5e9')
        else:
            shortfall = target_attendance - total_attendance
            conclusion_text = f"""
            <b>✗ KESIMPULAN: BELUM MENCUKUPI TARGET</b><br/>
            Pegawai <b>BELUM MEMENUHI</b> target kehadiran senam.<br/>
            Kurang <b>{shortfall} kali</b> dari target <b>{target_attendance} kali</b><br/>
            Persentase pencapaian: <b>{(total_attendance/target_attendance*100) if target_attendance > 0 else 0:.1f}%</b> dari target {percentage_target}%
            """
            conclusion_color = colors.HexColor('#c62828')
            bg_color = colors.HexColor('#ffebee')
        
        conclusion_style = ParagraphStyle(
            'Conclusion',
            parent=styles['Normal'],
            fontSize=7.5,  # Font kecil
            textColor=conclusion_color,
            spaceBefore=3,
            spaceAfter=5,
            fontName='Helvetica-Bold',
            leftIndent=6,
            rightIndent=6
        )
        
        conclusion_para = Paragraph(conclusion_text, conclusion_style)
        # Lebar: 14cm (muat A4)
        conclusion_table = Table([[conclusion_para]], colWidths=[14*cm])
        conclusion_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), bg_color),
            ('BOX', (0,0), (-1,-1), 1.5, conclusion_color),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        
        story.append(conclusion_table)
        story.append(Spacer(1, 0.25*cm))
        
        # TANDA TANGAN - TANPA GARIS
        signature_data = [
            ['', ''],
            ['KABAG SDM', 'KASUBAG KEPEGAWAIAN'],
            ['', ''],
            ['', ''],
            ['', ''],
            ['Dewi Nashrulloh, SKM.M.Kes', 'Rahmawati, SH'],
            ['NIK: 011205226', 'NIK: 022708322']
        ]
        
        # Lebar: 7+7 = 14cm
        signature_table = Table(signature_data, colWidths=[7*cm, 7*cm])
        signature_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 7),  # Font kecil
            ('FONTSIZE', (0,1), (-1,1), 7),
            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
            ('TOPPADDING', (0,1), (-1,1), 0),
            ('TOPPADDING', (0,5), (-1,6), 3),
        ]))
        
        story.append(signature_table)
        
        # FOOTER
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=6,
            alignment=TA_RIGHT,
            textColor=colors.grey
        )
        
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(
            f"Dicetak pada: {datetime.now().strftime('%d %B %Y %H:%M:%S')}",
            footer_style
        ))
        
        # BUILD PDF
        doc.build(story)
        
        buffer.seek(0)
        
        filename = f"rekap_senam_{employee_data.get('nik', 'unknown')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})

@app.route('/api/export-group-pdf', methods=['POST'])
def export_group_pdf():
    """Export PDF untuk kelompok pegawai - A4 PORTRAIT - FLEKSIBEL RENTANG WAKTU"""
    try:
        data = request.json
        employees = data.get('employees', [])
        date_range = data.get('date_range', {})
        shift_filter = data.get('shift_filter', 'all')
        struktur_lini = data.get('struktur_lini', 'Semua')
        
        print(f"DEBUG: Shift filter received: {shift_filter}")
        print(f"DEBUG: Total employees before filter: {len(employees)}")
        print(f"DEBUG: Date range: {date_range}")
        
        if not employees:
            return jsonify({'success': False, 'message': 'Tidak ada data pegawai'})
        
        # DEEP COPY employees untuk menghindari modifikasi data asli
        import copy
        employees = copy.deepcopy(employees)
        
        # Filter berdasarkan shift status jika bukan "all"
        if shift_filter != 'all':
            filtered_employees = []
            for emp in employees:
                if not isinstance(emp, dict):
                    print(f"WARNING: Invalid employee data type: {type(emp)}")
                    continue
                
                emp_shift_status = emp.get('shift_status', 'non_shift')
                print(f"DEBUG: Employee {emp.get('nama', 'UNKNOWN')} - shift_status: {emp_shift_status}")
                
                if emp_shift_status == shift_filter:
                    filtered_employees.append(emp)
            
            employees = filtered_employees
            print(f"DEBUG: Total employees after filter: {len(employees)}")
        
        if not employees:
            return jsonify({'success': False, 'message': 'Tidak ada data pegawai untuk filter shift yang dipilih'})
        
        # CREATE PDF - A4 PORTRAIT
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=1.2*cm,
            leftMargin=1.2*cm,
            topMargin=0.8*cm,
            bottomMargin=1.5*cm
        )
        
        story = []
        styles = getSampleStyleSheet()
        
        # CUSTOM STYLES
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=11,
            alignment=TA_CENTER,
            spaceAfter=3,
            textColor=colors.HexColor('#1a237e'),
            fontName='Helvetica-Bold'
        )
        
        subtitle_style = ParagraphStyle(
            'SubtitleStyle',
            parent=styles['Heading2'],
            fontSize=9,
            alignment=TA_CENTER,
            spaceAfter=6,
            textColor=colors.HexColor('#283593'),
            fontName='Helvetica-Bold'
        )
        
        # HEADER
        story.append(Paragraph("RUMAH SAKIT ISLAM SITI KHADIJAH PALEMBANG", title_style))
        story.append(Paragraph("REKAP ABSENSI SENAM PEGAWAI PER KELOMPOK", subtitle_style))
        story.append(Spacer(1, 0.2*cm))
        
        # Filter shift text
        shift_text_map = {
            'all': 'Semua (Shift & Non-Shift)',
            'shift': 'Shift',
            'non_shift': 'Non-Shift'
        }
        shift_text = shift_text_map.get(shift_filter, 'Semua')
        
        # INFO KELOMPOK - Compact
        info_data = [
            ['STRUKTUR LINI', ':', struktur_lini, 'FILTER SHIFT', ':', shift_text],
            ['JUMLAH PEGAWAI', ':', str(len(employees)), 'PERIODE', ':', f"{date_range.get('start', '-')} s/d {date_range.get('end', '-')}"]
        ]
        
        info_table = Table(info_data, colWidths=[2.8*cm, 0.3*cm, 4*cm, 2.8*cm, 0.3*cm, 4*cm])
        info_table.setStyle(TableStyle([
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (3,0), (3,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 7),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 2),
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 0.3*cm))
        
        # ===== PERBAIKAN: Ambil bulan dari rentang waktu FLEKSIBEL =====
        months_list = []
        if date_range.get('start') and date_range.get('end'):
            start_year, start_month = map(int, date_range['start'].split('-'))
            end_year, end_month = map(int, date_range['end'].split('-'))
            
            current_year = start_year
            current_month = start_month
            
            # Loop untuk generate semua bulan dalam rentang
            while (current_year < end_year) or (current_year == end_year and current_month <= end_month):
                month_key = f"{current_year}-{str(current_month).zfill(2)}"
                months_list.append(month_key)
                
                current_month += 1
                if current_month > 12:
                    current_month = 1
                    current_year += 1
        
        print(f"DEBUG: Total months generated: {len(months_list)}")
        print(f"DEBUG: Months list: {months_list[:5]}...{months_list[-5:] if len(months_list) > 5 else ''}")
        
        # TABEL DATA PEGAWAI - DINAMIS berdasarkan jumlah bulan
        month_names_short = {
            '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
            '05': 'Mei', '06': 'Jun', '07': 'Jul', '08': 'Ags',
            '09': 'Sep', '10': 'Okt', '11': 'Nov', '12': 'Des'
        }
        
        # ===== PERBAIKAN: Hitung lebar kolom dinamis =====
        num_months = len(months_list)
        
        # Tentukan ukuran font dan padding berdasarkan jumlah bulan
        if num_months <= 12:
            # 12 bulan atau kurang - ukuran normal
            header_font_size = 5
            data_font_size = 5.5
            month_col_width = 0.55*cm
            show_year_in_header = True
        elif num_months <= 18:
            # 13-18 bulan - ukuran sedang
            header_font_size = 4.5
            data_font_size = 5
            month_col_width = 0.50*cm
            show_year_in_header = True
        elif num_months <= 24:
            # 19-24 bulan - ukuran kecil
            header_font_size = 4
            data_font_size = 4.5
            month_col_width = 0.45*cm
            show_year_in_header = False  # Hanya bulan, tanpa tahun
        else:
            # Lebih dari 24 bulan - sangat kecil
            header_font_size = 3.5
            data_font_size = 4
            month_col_width = 0.40*cm
            show_year_in_header = False
        
        # Header tabel dengan 2 kolom status (Non-Shift dan Shift)
        header = ['NO', 'NAMA', 'NIK', 'TEMPAT TUGAS']
        for month_key in months_list:
            year, month = month_key.split('-')
            if show_year_in_header:
                header.append(f"{month_names_short[month]}\n'{year[2:]}")
            else:
                header.append(f"{month_names_short[month]}")
        header.extend(['TOTAL', 'NON-SHIFT', 'SHIFT'])
        
        table_data = [header]
        
        # ===== PERBAIKAN: Hitung target berdasarkan JUMLAH BULAN AKTUAL =====
        # Asumsi: 40 kegiatan per tahun = 40/12 per bulan ≈ 3.33 per bulan
        kegiatan_per_bulan = 40 / 12
        total_kegiatan_dalam_periode = kegiatan_per_bulan * num_months
        
        # Target untuk periode ini
        TARGET_NON_SHIFT = int(total_kegiatan_dalam_periode * 0.70)  # 70%
        TARGET_SHIFT = int(total_kegiatan_dalam_periode * 0.50)      # 50%
        
        print(f"DEBUG: Total kegiatan dalam periode {num_months} bulan: {total_kegiatan_dalam_periode:.2f}")
        print(f"DEBUG: Target NON-SHIFT: {TARGET_NON_SHIFT}")
        print(f"DEBUG: Target SHIFT: {TARGET_SHIFT}")
        
        # Data pegawai dengan status
        grand_total = 0
        for idx, emp in enumerate(employees, 1):
            try:
                if not isinstance(emp, dict):
                    print(f"ERROR: Invalid employee at index {idx}")
                    continue
                
                nama = emp.get('nama', f'Pegawai {idx}')
                nik = emp.get('nik', '-')
                tempat = emp.get('tempat', '-')
                
                # Gunakan Paragraph untuk text wrapping
                nama_para = Paragraph(str(nama), ParagraphStyle(
                    'NamaStyle',
                    fontSize=data_font_size,
                    leading=data_font_size + 1,
                    wordWrap='CJK'
                ))
                
                nik_para = Paragraph(str(nik), ParagraphStyle(
                    'NikStyle',
                    fontSize=data_font_size,
                    leading=data_font_size + 1,
                    wordWrap='CJK'
                ))
                
                tempat_para = Paragraph(str(tempat), ParagraphStyle(
                    'TempatStyle',
                    fontSize=data_font_size,
                    leading=data_font_size + 1,
                    wordWrap='CJK'
                ))
                
                row = [
                    str(idx),
                    nama_para,
                    nik_para,
                    tempat_para
                ]
                
                employee_total = 0
                for month_key in months_list:
                    year, month = month_key.split('-')
                    value = 0
                    
                    # Validasi struktur data bulanan
                    bulanan = emp.get('bulanan', {})
                    if not isinstance(bulanan, dict):
                        print(f"WARNING: Invalid bulanan data for {nama}")
                        row.append('0')
                        continue
                    
                    if year in bulanan:
                        year_data = bulanan[year]
                        if isinstance(year_data, dict):
                            month_data = year_data.get(month_key, {})
                            if isinstance(month_data, dict):
                                value = month_data.get('value', 0)
                                try:
                                    value = int(value) if value else 0
                                except (ValueError, TypeError):
                                    value = 0
                    
                    row.append(str(value))
                    employee_total += value
                
                row.append(str(employee_total))
                grand_total += employee_total
                
                # Status NON-SHIFT (target dinamis)
                if employee_total >= TARGET_NON_SHIFT:
                    status_non_shift = Paragraph('✓ Tercapai', ParagraphStyle(
                        'StatusGood',
                        fontSize=data_font_size,
                        textColor=colors.HexColor('#2e7d32'),
                        fontName='Helvetica-Bold',
                        alignment=TA_CENTER
                    ))
                else:
                    kurang_non_shift = TARGET_NON_SHIFT - employee_total
                    status_non_shift = Paragraph(f'✗ Kurang {kurang_non_shift}x', ParagraphStyle(
                        'StatusBad',
                        fontSize=data_font_size,
                        textColor=colors.HexColor('#c62828'),
                        alignment=TA_CENTER
                    ))
                
                # Status SHIFT (target dinamis)
                if employee_total >= TARGET_SHIFT:
                    status_shift = Paragraph('✓ Tercapai', ParagraphStyle(
                        'StatusGood',
                        fontSize=data_font_size,
                        textColor=colors.HexColor('#2e7d32'),
                        fontName='Helvetica-Bold',
                        alignment=TA_CENTER
                    ))
                else:
                    kurang_shift = TARGET_SHIFT - employee_total
                    status_shift = Paragraph(f'✗ Kurang {kurang_shift}x', ParagraphStyle(
                        'StatusBad',
                        fontSize=data_font_size,
                        textColor=colors.HexColor('#c62828'),
                        alignment=TA_CENTER
                    ))
                
                row.extend([status_non_shift, status_shift])
                table_data.append(row)
                
            except Exception as e:
                print(f"ERROR processing employee {idx}: {str(e)}")
                import traceback
                traceback.print_exc()
                continue
        
        # Total row
        total_row = ['', Paragraph('<b>TOTAL</b>', ParagraphStyle(
            'TotalBold',
            fontSize=6,
            fontName='Helvetica-Bold'
        )), '', '']
        
        for month_key in months_list:
            month_total = 0
            try:
                for emp in employees:
                    if isinstance(emp, dict):
                        bulanan = emp.get('bulanan', {})
                        if isinstance(bulanan, dict):
                            year = month_key.split('-')[0]
                            if year in bulanan:
                                year_data = bulanan[year]
                                if isinstance(year_data, dict):
                                    month_data = year_data.get(month_key, {})
                                    if isinstance(month_data, dict):
                                        value = month_data.get('value', 0)
                                        try:
                                            month_total += int(value) if value else 0
                                        except (ValueError, TypeError):
                                            pass
            except Exception as e:
                print(f"ERROR calculating month total for {month_key}: {str(e)}")
            
            total_row.append(str(month_total))
        
        total_row.extend([str(grand_total), '', ''])
        table_data.append(total_row)
        
        # Validasi table_data
        if len(table_data) <= 1:
            return jsonify({'success': False, 'message': 'Tidak ada data valid untuk ditampilkan'})
        
        # ===== PERBAIKAN: Hitung lebar kolom DINAMIS =====
        # Total width: 18cm - NO(0.5) - NAMA(2.5) - NIK(1.5) - TEMPAT(2.5) - TOTAL(0.8) - NON-SHIFT(1.6) - SHIFT(1.6) = sisa untuk bulan
        fixed_width = 0.5 + 2.5 + 1.5 + 2.5 + 0.8 + 1.6 + 1.6  # = 11.0cm
        available_for_months = 18 - fixed_width  # = 7cm
        
        # Hitung lebar kolom bulan
        if num_months > 0:
            month_col_width = min(month_col_width, available_for_months / num_months)
        
        col_widths = [
            0.5*cm,    # NO
            2.5*cm,    # NAMA
            1.5*cm,    # NIK
            2.5*cm,    # TEMPAT TUGAS
        ] + [month_col_width*cm] * num_months + [
            0.8*cm,    # TOTAL
            1.6*cm,    # NON-SHIFT
            1.6*cm     # SHIFT
        ]
        
        data_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        data_table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1976d2')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), header_font_size),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
            
            # Data rows
            ('FONTSIZE', (0,1), (-1,-2), data_font_size),
            ('ALIGN', (0,1), (0,-1), 'CENTER'),  # NO
            ('ALIGN', (2,1), (2,-2), 'CENTER'),  # NIK
            ('ALIGN', (4,1), (-1,-2), 'CENTER'),  # Bulan, Total, Status
            ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),
            
            # Total row
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#e3f2fd')),
            ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,-1), (-1,-1), 6),
            ('ALIGN', (4,-1), (-1,-1), 'CENTER'),
            
            # Grid
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#f5f5f5')]),
            
            # Padding
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 1.5),
            ('RIGHTPADDING', (0,0), (-1,-1), 1.5),
        ]))
        
        story.append(data_table)
        story.append(Spacer(1, 0.4*cm))
        
        # TARGET DAN ANALISIS - DINAMIS
        avg_attendance = grand_total / len(employees) if len(employees) > 0 else 0
        
        analysis_data = [
            ['KETERANGAN', 'NILAI'],
            ['Total Pegawai', f'{len(employees)} orang'],
            ['Periode', f'{num_months} bulan ({date_range.get("start")} s/d {date_range.get("end")})'],
            ['Total Kehadiran', f'{grand_total} kali'],
            ['Rata-rata per Pegawai', f'{avg_attendance:.1f} kali'],
        ]
        
        analysis_table = Table(analysis_data, colWidths=[10*cm, 5*cm])
        analysis_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e3f2fd')),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 7),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (1,0), (1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.grey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#fafafa')]),
        ]))
        
        story.append(analysis_table)
        story.append(Spacer(1, 0.25*cm))
        
        # KETERANGAN TARGET - DINAMIS
        keterangan_style = ParagraphStyle(
            'Keterangan',
            parent=styles['Normal'],
            fontSize=6.5,
            leading=8,
            leftIndent=8,
            rightIndent=8,
            spaceAfter=2
        )
        
        keterangan_text = f"""
        <b>KETERANGAN STATUS:</b><br/>
        • <b>Periode:</b> {num_months} bulan dari {date_range.get('start')} sampai {date_range.get('end')}<br/>
        • <b>Total Kegiatan Senam dalam periode:</b> {total_kegiatan_dalam_periode:.1f} kali ({kegiatan_per_bulan:.2f} kali/bulan × {num_months} bulan)<br/>
        • <b>Target NON-SHIFT:</b> {TARGET_NON_SHIFT} kali (70% dari {total_kegiatan_dalam_periode:.1f})<br/>
        • <b>Target SHIFT:</b> {TARGET_SHIFT} kali (50% dari {total_kegiatan_dalam_periode:.1f})<br/>
        • Status <font color="#2e7d32"><b>✓ Tercapai</b></font> = Kehadiran mencapai/melebihi target<br/>
        • Status <font color="#c62828"><b>✗ Kurang Nx</b></font> = Kehadiran kurang N kali dari target
        """
        
        keterangan_para = Paragraph(keterangan_text, keterangan_style)
        keterangan_table = Table([[keterangan_para]], colWidths=[15*cm])
        keterangan_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fff9c4')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#f9a825')),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        
        story.append(keterangan_table)
        story.append(Spacer(1, 0.4*cm))
        
        # TANDA TANGAN
        signature_data = [
            ['', ''],
            ['KABAG SDM', 'KASUBAG KEPEGAWAIAN'],
            ['', ''],
            ['', ''],
            ['Dewi Nashrulloh, SKM.M.Kes', 'Rahmawati, SH'],
            ['NIK: 011205226', 'NIK: 022708322']
        ]
        
        signature_table = Table(signature_data, colWidths=[7.5*cm, 7.5*cm])
        signature_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 7),
            ('FONTSIZE', (0,1), (-1,1), 7),
            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
            ('TOPPADDING', (0,1), (-1,1), 0),
            ('TOPPADDING', (0,4), (-1,5), 3),
        ]))
        
        story.append(signature_table)
        
        # FOOTER
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=6,
            alignment=TA_RIGHT,
            textColor=colors.grey
        )
        
        story.append(Spacer(1, 0.15*cm))
        story.append(Paragraph(
            f"Dicetak pada: {datetime.now().strftime('%d %B %Y %H:%M:%S')}",
            footer_style
        ))
        
        # BUILD PDF
        try:
            print(f"DEBUG: Building PDF with {len(story)} elements")
            doc.build(story)
            print("DEBUG: PDF build successful")
        except Exception as build_error:
            print(f"ERROR building PDF: {str(build_error)}")
            import traceback
            traceback.print_exc()
            return jsonify({'success': False, 'message': f'Gagal membuat PDF: {str(build_error)}'})
        
        # Validasi buffer
        buffer.seek(0)
        pdf_size = len(buffer.getvalue())
        print(f"DEBUG: PDF size: {pdf_size} bytes")
        
        if pdf_size < 100:
            return jsonify({'success': False, 'message': 'PDF yang dihasilkan tidak valid (terlalu kecil)'})
        
        struktur_safe = struktur_lini.replace(' ', '_').replace('/', '-')
        shift_safe = shift_filter.replace('_', '-')
        filename = f"rekap_senam_kelompok_{struktur_safe}_{shift_safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        print(f"DEBUG: Returning PDF file: {filename}, size: {pdf_size}")
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"Group PDF Export error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})
    

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    try:
        data = request.json.get('data', [])
        
        if not data:
            return jsonify({'success': False, 'message': 'Tidak ada data'})
        
        rows = []
        for pegawai in data:
            row = {
                'NAMA': pegawai['nama'],
                'NIK': pegawai['nik'],
                'JENIS_KELAMIN': pegawai['jk'],
                'STATUS_PEGAWAI': pegawai['status'],
                'KELOMPOK_NAKES': pegawai['kelompok'],
                'JABATAN': pegawai['jabatan'],
                'STRUKTUR_LINI': pegawai['struktur'],
                'TEMPAT_TUGAS': pegawai['tempat'],
                'TOTAL_SEMUA_TAHUN': pegawai['total_all']
            }
            
            for year in sorted(pegawai['tahunan'].keys()):
                row[f'TOTAL_{year}'] = pegawai['tahunan'][year]
            
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekap Senam')
            
            worksheet = writer.sheets['Rekap Senam']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        filename = f"rekap_senam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export error: {str(e)}")
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})

@app.route('/api/export-group-excel', methods=['POST'])
def export_group_excel():
    """Export Excel untuk kelompok pegawai (berdasarkan filter struktur lini)"""
    try:
        data = request.json
        employees = data.get('employees', [])
        date_range = data.get('date_range', {})
        struktur_lini = data.get('struktur_lini', 'Semua')
        
        if not employees:
            return jsonify({'success': False, 'message': 'Tidak ada data pegawai'})
        
        # Ambil bulan dari rentang waktu
        months_list = []
        if date_range.get('start') and date_range.get('end'):
            start_year, start_month = map(int, date_range['start'].split('-'))
            end_year, end_month = map(int, date_range['end'].split('-'))
            
            current_year = start_year
            current_month = start_month
            
            while (current_year < end_year) or (current_year == end_year and current_month <= end_month):
                month_key = f"{current_year}-{str(current_month).zfill(2)}"
                months_list.append(month_key)
                
                current_month += 1
                if current_month > 12:
                    current_month = 1
                    current_year += 1
        
        # Buat data untuk Excel
        rows = []
        for emp in employees:
            row = {
                'NAMA': emp.get('nama', '-'),
                'NIK': emp.get('nik', '-'),
                'JABATAN': emp.get('jabatan', '-'),
                'STRUKTUR_LINI': emp.get('struktur', '-'),
            }
            
            employee_total = 0
            for month_key in months_list:
                year, month = month_key.split('-')
                value = 0
                
                if year in emp.get('bulanan', {}):
                    month_data = emp['bulanan'][year].get(month_key, {})
                    value = month_data.get('value', 0)
                
                row[month_key] = value
                employee_total += value
            
            row['TOTAL'] = employee_total
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekap Senam Kelompok')
            
            worksheet = writer.sheets['Rekap Senam Kelompok']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        struktur_safe = struktur_lini.replace(' ', '_').replace('/', '-')
        filename = f"rekap_senam_kelompok_{struktur_safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Group Excel Export error: {str(e)}")
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})

@app.route('/api/validate-template', methods=['POST'])
def validate_template():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'valid': False, 'message': 'Tidak ada file'})
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'valid': False, 'message': 'Nama file kosong'})
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'valid': False, 'message': 'Format tidak didukung'})
        
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(temp_path)
        
        try:
            if temp_path.endswith('.csv'):
                df = pd.read_csv(temp_path, header=4, nrows=5, encoding='utf-8')
            else:
                df = pd.read_excel(temp_path, header=4, nrows=5)
            
            df.columns = [normalize_col(c) for c in df.columns]
            
            required_columns = ['NAMA', 'NIK']
            missing_columns = []
            
            for col in required_columns:
                if col not in df.columns:
                    missing_columns.append(col)
            
            if temp_path.endswith('.csv'):
                df_full = pd.read_csv(temp_path, header=4, encoding='utf-8')
            else:
                df_full = pd.read_excel(temp_path, header=4)
            
            data_rows = len(df_full.dropna(subset=['NAMA'] if 'NAMA' in df_full.columns else []))
            
            shutil.rmtree(temp_dir)
            
            if missing_columns:
                return jsonify({
                    'success': True,
                    'valid': False,
                    'message': f'Kolom wajib tidak ditemukan: {", ".join(missing_columns)}',
                    'missing_columns': missing_columns,
                    'data_rows': data_rows
                })
            
            columns_found = list(df.columns)[:15] if len(df.columns) > 0 else []
            
            return jsonify({
                'success': True,
                'valid': True,
                'message': 'Format file valid',
                'data_rows': data_rows,
                'columns_found': columns_found[:10],
                'total_columns': len(df.columns)
            })
            
        except Exception as e:
            shutil.rmtree(temp_dir)
            return jsonify({
                'success': False,
                'valid': False,
                'message': f'Error membaca file: {str(e)}'
            })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'valid': False,
            'message': f'Error validasi: {str(e)}'
        })

@app.route('/api/clear-data', methods=['POST'])
def clear_data():
    try:
        clear_temp_data()
        return jsonify({'success': True, 'message': 'Data berhasil dihapus'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'message': 'Server is running'})

if __name__ == '__main__':
    clear_temp_data()
    
    print("=" * 60)
    print("DASHBOARD REKAP SENAM - FIXED VERSION")
    print("=" * 60)
    print("✓ Export PDF Kelompok: A4 Portrait")
    print("✓ Tanda tangan tidak terpotong")
    print("✓ Filter shift diperbaiki (all, shift, non_shift)")
    print("=" * 60)
    print("Server: http://localhost:5001")
    print("=" * 60)
    
    app.run(debug=True, port=5001, host='0.0.0.0')

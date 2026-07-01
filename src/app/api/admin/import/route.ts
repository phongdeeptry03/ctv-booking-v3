import { NextRequest, NextResponse } from 'next/server';
import { verifyAdminSession } from '@/lib/auth';
import { appendTaskRows } from '@/lib/sheets';
import { syncTaskLimit, ensureSchema } from '@/lib/db';
import * as XLSX from 'xlsx';

export const dynamic = 'force-dynamic';

const HEADER_ALIASES: Record<string, string> = {
  'ngay': 'Ngay',
  'team': 'Team',
  'task': 'Task',
  'giobatdau': 'GioBatDau',
  'thoigianbatdau': 'GioBatDau',
  'gioketthuc': 'GioKetThuc',
  'thoigianketthuc': 'GioKetThuc',
  'gionghi': 'GioNghi',
  'thoigiannghigiaolao': 'GioNghi',
  'thoigiannghigiailao': 'GioNghi',
  'nghigiailao': 'GioNghi',
  'loainhanvien': 'LoaiNhanVien',
  'loainv': 'LoaiNhanVien',
  'sobookbpopt': 'SoLuongCan',
  'soluongcan': 'SoLuongCan',
  'soluong': 'SoLuongCan',
};

function normalizeKey(s: string): string {
  return s.toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '');
}

/** Chuyển bất kỳ giá trị cell nào thành string ngày YYYY-MM-DD */
function toDateStr(val: any): string {
  if (!val) return '';
  // Date object (cellDates: true)
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(val).trim();
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // M/D/YYYY hoac D/M/YYYY - xu ly ca hai, uu tien MM/DD/YYYY (Excel My)
  const parts = s.split('/');
  if (parts.length === 3) {
    const [a, b, y] = parts;
    // Excel thuong dung M/D/YYYY
    const mm = a.padStart(2, '0');
    const dd = b.padStart(2, '0');
    return `${y}-${mm}-${dd}`;
  }
  return s;
}

/** Chuyển bất kỳ giá trị cell nào thành string giờ H:mm */
function toTimeStr(val: any): string {
  if (!val && val !== 0) return '';
  // Date object tu cellDates:true (gio duoc bieu dien la Date 1899-12-30 + fraction)
  if (val instanceof Date) {
    const h = val.getHours();
    const m = String(val.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  const s = String(val).trim();
  // H:mm hoac HH:mm hoac H:mm:ss
  if (/^\d{1,2}:\d{2}/.test(s)) {
    const parts = s.split(':');
    return `${parseInt(parts[0])}:${parts[1]}`;
  }
  // So thap phan (0.75 = 18:00)
  const num = parseFloat(s);
  if (!isNaN(num) && num >= 0 && num <= 1) {
    const totalMin = Math.round(num * 24 * 60);
    const h = Math.floor(totalMin / 60) % 24;
    const m = totalMin % 60;
    return `${h}:${String(m).padStart(2, '0')}`;
  }
  // "6:00 PM" format
  const pmMatch = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (pmMatch) {
    let h = parseInt(pmMatch[1]);
    const m = pmMatch[2];
    if (pmMatch[3].toUpperCase() === 'PM' && h !== 12) h += 12;
    if (pmMatch[3].toUpperCase() === 'AM' && h === 12) h = 0;
    return `${h}:${m}`;
  }
  return s;
}

export async function POST(req: NextRequest) {
  if (!(await verifyAdminSession())) {
    return NextResponse.json({ error: 'Chua dang nhap' }, { status: 401 });
  }

  try {
    await ensureSchema();

    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) {
      return NextResponse.json({ error: 'Vui long chon file' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    // cellDates: true -> xlsx tu chuyen serial number thanh JS Date object
    // raw: true -> giu nguyen gia tri goc, khong ep sang string
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
    });

    if (rows.length < 2) {
      return NextResponse.json({ error: 'File khong co du lieu' }, { status: 400 });
    }

    // Map header -> field
    const headerRow = rows[0].map((h: any) => String(h ?? ''));
    const colMap: Record<number, string> = {};
    headerRow.forEach((h: string, idx: number) => {
      const key = normalizeKey(h);
      if (HEADER_ALIASES[key]) colMap[idx] = HEADER_ALIASES[key];
    });

    // Kiem tra cot bat buoc
    const required = ['Ngay', 'Team', 'Task', 'GioBatDau', 'GioKetThuc', 'SoLuongCan'];
    const mapped = new Set(Object.values(colMap));
    const missing = required.filter(f => !mapped.has(f));
    if (missing.length > 0) {
      const actualCols = headerRow.map((h: string) => `"${h}"(key:${normalizeKey(h)})`).join(', ');
      return NextResponse.json({
        error: `Khong tim thay cot: ${missing.join(', ')}. Cot trong file: ${actualCols}`,
      }, { status: 400 });
    }

    // Parse tung dong
    const tasksToImport: any[] = [];
    for (const row of rows.slice(1)) {
      if (!row || row.every((c: any) => c === null || c === '')) continue;

      const get = (field: string) => {
        const idx = Object.entries(colMap).find(([, f]) => f === field)?.[0];
        return idx !== undefined ? row[Number(idx)] : null;
      };

      const ngay = toDateStr(get('Ngay'));
      const team = String(get('Team') ?? '').trim();
      const task = String(get('Task') ?? '').trim();
      if (!ngay || !team || !task) continue;

      const id = `T${Date.now()}${Math.floor(Math.random() * 100000)}`;
      tasksToImport.push({
        ID: id,
        Ngay: ngay,
        Team: team,
        Task: task,
        GioBatDau: toTimeStr(get('GioBatDau')),
        GioKetThuc: toTimeStr(get('GioKetThuc')),
        GioNghi: toTimeStr(get('GioNghi')),
        LoaiNhanVien: String(get('LoaiNhanVien') ?? '').trim(),
        SoLuongCan: Number(get('SoLuongCan')) || 0,
        DaDangKy: 0,
      });
    }

    if (tasksToImport.length === 0) {
      return NextResponse.json({ error: 'Khong co dong du lieu hop le' }, { status: 400 });
    }

    // Log de debug (xem trong Vercel Functions logs)
    console.log('Import preview:', JSON.stringify(tasksToImport.slice(0, 2), null, 2));

    await appendTaskRows(tasksToImport);
    for (const t of tasksToImport) {
      await syncTaskLimit(t.ID, t.SoLuongCan);
    }

    return NextResponse.json({
      success: true,
      imported: tasksToImport.length,
      message: `Da import ${tasksToImport.length} task thanh cong`,
      // Tra ve preview de debug
      preview: tasksToImport.slice(0, 3).map(t => ({ Ngay: t.Ngay, GioBatDau: t.GioBatDau, Task: t.Task })),
    });
  } catch (err: any) {
    console.error('Import error:', err);
    return NextResponse.json({ error: err?.message || 'Loi he thong' }, { status: 500 });
  }
}

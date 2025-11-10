<?php

namespace App\Imports;

use App\Models\VehicleRegistration;
use Carbon\Carbon;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithHeadingRow;
use Maatwebsite\Excel\Concerns\WithChunkReading;
use Maatwebsite\Excel\Concerns\WithBatchInserts;
use Maatwebsite\Excel\Concerns\SkipsEmptyRows;

class VehicleRegistrationsImport implements
    ToModel,
    WithHeadingRow,
    WithChunkReading,
    WithBatchInserts,
    SkipsEmptyRows
{
    /**
    * @param array $row
    *
    * @return \Illuminate\Database\Eloquent\Model|null
    */
    public function model(array $row)
    {
        // Bỏ qua dòng tiêu đề hoặc dòng trống
        if (empty($row['sst']) || $row['sst'] === 'SST') {
            return null;
        }

        // Xử lý STT - bỏ qua nếu là công thức Excel hoặc không phải số
        $stt = $row['sst'] ?? null;

        // Kiểm tra nếu là công thức Excel (bắt đầu bằng =)
        if (is_string($stt) && strpos($stt, '=') === 0) {
            return null; // Bỏ qua dòng có công thức
        }

        // Chuyển đổi sang số nguyên
        $stt = filter_var($stt, FILTER_VALIDATE_INT);

        // Bỏ qua nếu không phải số hợp lệ
        if ($stt === false || $stt === null) {
            return null;
        }

        // Lấy ngày đăng ký (nếu có)
        $ngayDangKy = null;
        $thang = null;
        $nam = null;

        if (!empty($row['ngay_dang_ky'])) {
            try {
                // Xử lý ngày từ Excel
                if (is_numeric($row['ngay_dang_ky'])) {
                    // Excel date serial number
                    $ngayDangKy = Carbon::instance(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($row['ngay_dang_ky']));
                } else {
                    // String date
                    $ngayDangKy = Carbon::parse($row['ngay_dang_ky']);
                }

                $thang = $ngayDangKy->month;
                $nam = $ngayDangKy->year;
            } catch (\Exception $e) {
                // Nếu không parse được ngày, dùng tháng năm hiện tại
                $ngayDangKy = null;
            }
        }

        // Nếu không có ngày đăng ký, dùng tháng năm hiện tại
        if (!$thang || !$nam) {
            $thang = now()->month;
            $nam = now()->year;
        }

        return new VehicleRegistration([
            'stt' => $stt,
            'ho_va_ten' => $row['ho_va_ten'] ?? '',
            'lop' => $row['lop'] ?? '',
            'loai_xe' => $row['loai_xe'] ?? '',
            'bien_so' => $row['bien_so'] ?? null,
            'so_ve_xe' => $row['so_ve_xe'] ?? '',
            'ngay_dang_ky' => $ngayDangKy,
            'thang' => $thang,
            'nam' => $nam,
        ]);
    }

    /**
     * Chunk size để xử lý từng phần
     */
    public function chunkSize(): int
    {
        return 100;
    }

    /**
     * Batch insert để tăng hiệu suất
     */
    public function batchSize(): int
    {
        return 100;
    }

    /**
     * Chỉ định dòng nào là heading
     */
    public function headingRow(): int
    {
        return 2; // Dòng thứ 2 là header (dòng 1 là tiêu đề)
    }
}

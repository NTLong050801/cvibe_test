<?php

namespace App\Http\Controllers;

use App\Models\VehicleRegistration;
use Illuminate\Http\Request;

class VehicleRegistrationController extends Controller
{
    public function index(Request $request)
    {
        $query = VehicleRegistration::query();

        // Lọc theo tháng năm nếu có
        if ($request->has('thang') && $request->thang) {
            $query->where('thang', $request->thang);
        }
        if ($request->has('nam') && $request->nam) {
            $query->where('nam', $request->nam);
        }

        // Phân trang
        $registrations = $query->orderBy('created_at', 'desc')->paginate(15);

        return view('vehicle-registrations.index', compact('registrations'));
    }

    public function import(Request $request)
    {
        $request->validate([
            'file' => 'required|mimes:xlsx,xls|max:10240', // Max 10MB
        ]);

        try {
            $file = $request->file('file');
            $fileName = time() . '_' . $file->getClientOriginalName();
            $filePath = $file->storeAs('imports', $fileName);

            // Dispatch job để xử lý import
            \App\Jobs\ImportVehicleRegistrationsJob::dispatch(storage_path('app/' . $filePath));

            return back()->with('success', 'File đã được upload! Đang xử lý import trong background...');
        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra: ' . $e->getMessage());
        }
    }

    public function getAvailableMonths()
    {
        $data = VehicleRegistration::selectRaw('DISTINCT nam, thang')
            ->orderBy('nam', 'desc')
            ->orderBy('thang', 'desc')
            ->get()
            ->groupBy('nam')
            ->map(function ($items) {
                return $items->pluck('thang')->toArray();
            });

        return response()->json($data);
    }

    public function export(Request $request)
    {
        $request->validate([
            'thang' => 'required|integer|min:1|max:12',
            'nam' => 'required|integer|min:2020|max:2030',
        ]);

        $thang = $request->thang;
        $nam = $request->nam;

        // Kiểm tra xem có dữ liệu không
        $count = VehicleRegistration::where('thang', $thang)
            ->where('nam', $nam)
            ->count();

        if ($count === 0) {
            return back()->with('error', "Không có dữ liệu cho tháng {$thang}/{$nam}!");
        }

        try {
            // Export trực tiếp và tải về ngay
            $templatePath = storage_path('Vexe_Template.xlsx');

            if (!file_exists($templatePath)) {
                return back()->with('error', 'Template file không tồn tại!');
            }

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($templatePath);
            $sheet = $spreadsheet->getActiveSheet();

            // Lấy dữ liệu từ database
            $registrations = VehicleRegistration::where('thang', $thang)
                ->where('nam', $nam)
                ->orderBy('stt')
                ->get();

            // Tính số nhóm cần thiết (mỗi nhóm có 3 vé)
            $totalTickets = $registrations->count();
            $groupsNeeded = ceil($totalTickets / 3);

            // TÌM VỊ TRÍ THỰC TẾ của các nhóm trong template
            $groupPositions = [];
            $highestRow = $sheet->getHighestRow();

            for ($row = 1; $row <= $highestRow; $row++) {
                $value = $sheet->getCell('A' . $row)->getValue();
                if ($value === 'VÉ GỬI XE') {
                    $groupPositions[] = $row;
                }
            }

            $templateGroupCount = count($groupPositions);

            // Nếu template không đủ nhóm, DUPLICATE thêm
            if ($templateGroupCount < $groupsNeeded) {
                $sourceGroupIndex = ($templateGroupCount > 1) ? 1 : 0;
                $sourceGroupStart = $groupPositions[$sourceGroupIndex];

                $groupHeight = ($templateGroupCount > 1)
                    ? ($groupPositions[1] - $groupPositions[0])
                    : 9;

                $lastGroupEnd = $groupPositions[$templateGroupCount - 1] + $groupHeight - 1;

                // Lưu thông tin ảnh từ nhóm mẫu
                $sourceImages = [];
                foreach ($sheet->getDrawingCollection() as $drawing) {
                    $coords = $drawing->getCoordinates();
                    preg_match('/([A-Z]+)(\d+)/', $coords, $matches);

                    if ($matches) {
                        $row = (int)$matches[2];
                        if ($row >= $sourceGroupStart && $row < ($sourceGroupStart + $groupHeight)) {
                            $sourceImages[] = [
                                'drawing' => $drawing,
                                'offset' => $row - $sourceGroupStart,
                            ];
                        }
                    }
                }

                // Lưu merged cells vào array TRƯỚC KHI iterate
                $mergedCellsArray = $sheet->getMergeCells();

                // Copy merged cells TRƯỚC để giữ nguyên structure
                $mergedCellsToAdd = [];
                foreach ($mergedCellsArray as $mergeCell) {
                    if (preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $mergeCell, $matches)) {
                        $startCol = $matches[1];
                        $startRow = (int)$matches[2];
                        $endCol = $matches[3];
                        $endRow = (int)$matches[4];

                        if ($startRow >= $sourceGroupStart && $startRow < $sourceGroupStart + $groupHeight) {
                            $rowOffset = $startRow - $sourceGroupStart;
                            $rowEndOffset = $endRow - $sourceGroupStart;

                            for ($groupIndex = $templateGroupCount; $groupIndex < $groupsNeeded; $groupIndex++) {
                                $newGroupStart = $lastGroupEnd + 1 + (($groupIndex - $templateGroupCount) * $groupHeight);
                                $newStartRow = $newGroupStart + $rowOffset;
                                $newEndRow = $newGroupStart + $rowEndOffset;
                                $mergedCellsToAdd[] = $startCol . $newStartRow . ':' . $endCol . $newEndRow;
                            }
                        }
                    }
                }

                // Duplicate thêm các nhóm còn thiếu
                for ($groupIndex = $templateGroupCount; $groupIndex < $groupsNeeded; $groupIndex++) {
                    $newGroupStart = $lastGroupEnd + 1 + (($groupIndex - $templateGroupCount) * $groupHeight);

                    // Copy từng dòng từ nhóm mẫu
                    for ($rowOffset = 0; $rowOffset < $groupHeight; $rowOffset++) {
                        $sourceRow = $sourceGroupStart + $rowOffset;
                        $targetRow = $newGroupStart + $rowOffset;

                        // Copy row height
                        $sheet->getRowDimension($targetRow)
                            ->setRowHeight($sheet->getRowDimension($sourceRow)->getRowHeight());

                        // Copy từng cell với style đầy đủ
                        foreach (range('A', 'I') as $col) {
                            $sourceCellCoord = $col . $sourceRow;
                            $targetCellCoord = $col . $targetRow;

                            // Copy value (chỉ label, không copy data B, E, H)
                            if (!in_array($col, ['B', 'E', 'H'])) {
                                $sheet->setCellValue($targetCellCoord, $sheet->getCell($sourceCellCoord)->getValue());
                            }

                            // Copy style bằng duplicateStyle
                            $sheet->duplicateStyle(
                                $sheet->getStyle($sourceCellCoord),
                                $targetCellCoord
                            );
                        }
                    }

                    // Copy column width (chỉ cần copy 1 lần)
                    if ($groupIndex == $templateGroupCount) {
                        foreach (range('A', 'I') as $col) {
                            $sheet->getColumnDimension($col)
                                ->setWidth($sheet->getColumnDimension($col)->getWidth());
                        }
                    }

                    // Clone ảnh cho nhóm mới
                    foreach ($sourceImages as $imageInfo) {
                        $sourceDrawing = $imageInfo['drawing'];
                        $offset = $imageInfo['offset'];

                        $newDrawing = clone $sourceDrawing;
                        $newRow = $newGroupStart + $offset;

                        preg_match('/([A-Z]+)(\d+)/', $sourceDrawing->getCoordinates(), $matches);
                        $col = $matches[1];

                        $newDrawing->setCoordinates($col . $newRow);
                        $newDrawing->setWorksheet($sheet);
                    }
                }

                // Merge cells cho các nhóm mới
                foreach ($mergedCellsToAdd as $mergeRange) {
                    try {
                        $sheet->mergeCells($mergeRange);
                    } catch (\Exception $e) {
                        // Ignore nếu đã merge hoặc conflict
                    }
                }

                // Cập nhật lại danh sách vị trí nhóm
                $groupPositions = [];
                for ($row = 1; $row <= $sheet->getHighestRow(); $row++) {
                    $value = $sheet->getCell('A' . $row)->getValue();
                    if ($value === 'VÉ GỬI XE') {
                        $groupPositions[] = $row;
                    }
                }
            }

            // ĐIỀN DỮ LIỆU vào các vé
            $ticketIndex = 0;
            foreach ($registrations as $registration) {
                $groupIndex = floor($ticketIndex / 3);
                $positionInGroup = $ticketIndex % 3;

                if ($groupIndex >= count($groupPositions)) {
                    break;
                }

                $groupStartRow = $groupPositions[$groupIndex];
                $colOffset = $positionInGroup * 3;
                $dataCol = chr(66 + $colOffset);

                // Điền dữ liệu vào cột tương ứng (B, E, hoặc H)
                $sheet->setCellValue($dataCol . ($groupStartRow + 1), $registration->so_ve_xe ?? '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 2), $registration->ho_va_ten ?? '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 3), $registration->lop ?? '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 4), $registration->bien_so ?? '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 5), $registration->loai_xe ?? '');

                $ticketIndex++;
            }

            // Xóa data từ các vé template không dùng
            for ($i = $ticketIndex; $i < $templateGroupCount * 3; $i++) {
                $groupIndex = floor($i / 3);
                $positionInGroup = $i % 3;

                if ($groupIndex >= count($groupPositions)) {
                    break;
                }

                $groupStartRow = $groupPositions[$groupIndex];
                $colOffset = $positionInGroup * 3;
                $dataCol = chr(66 + $colOffset);

                $sheet->setCellValue($dataCol . ($groupStartRow + 1), '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 2), '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 3), '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 4), '');
                $sheet->setCellValue($dataCol . ($groupStartRow + 5), '');
            }

            // Tạo file Excel và tải về
            $fileName = "Vexe_T{$thang}_{$nam}.xlsx";
            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');

            // Tạo response để tải file
            return response()->streamDownload(function() use ($writer) {
                $writer->save('php://output');
            }, $fileName, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ]);

        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra khi export: ' . $e->getMessage());
        }
    }

    public function deleteAll()
    {
        try {
            $count = VehicleRegistration::count();
            VehicleRegistration::truncate();

            return redirect()->route('vehicle-registrations.index')
                ->with('success', "Đã xóa thành công {$count} bản ghi!");
        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra khi xóa dữ liệu: ' . $e->getMessage());
        }
    }
}

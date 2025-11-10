<?php

namespace App\Jobs;

use App\Models\VehicleRegistration;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExportVehicleRegistrationsJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    protected $thang;
    protected $nam;

    /**
     * Số lần thử lại nếu job thất bại
     */
    public $tries = 3;

    /**
     * Timeout cho job (giây)
     */
    public $timeout = 600;

    /**
     * Create a new job instance.
     */
    public function __construct($thang, $nam)
    {
        $this->thang = $thang;
        $this->nam = $nam;
    }

    /**
     * Execute the job.
     */
    public function handle(): void
    {
        // Load template Excel
        $templatePath = storage_path('Vexe_Template.xlsx');

        if (!file_exists($templatePath)) {
            throw new \Exception('Template file not found: ' . $templatePath);
        }

        $spreadsheet = IOFactory::load($templatePath);
        $sheet = $spreadsheet->getActiveSheet();

        // Lấy dữ liệu từ database
        $registrations = VehicleRegistration::where('thang', $this->thang)
            ->where('nam', $this->nam)
            ->orderBy('stt')
            ->get();

        if ($registrations->isEmpty()) {
            return;
        }

        // Tính số nhóm cần thiết (mỗi nhóm có 3 vé)
        $totalTickets = $registrations->count();
        $groupsNeeded = ceil($totalTickets / 3);

        // TÌM VỊ TRÍ THỰC TẾ của các nhóm trong template (tự động phát hiện vị trí "VÉ GỬI XE")
        $groupPositions = [];
        $highestRow = $sheet->getHighestRow();

        for ($row = 1; $row <= $highestRow; $row++) {
            $value = $sheet->getCell('A' . $row)->getValue();
            if ($value === 'VÉ GỬI XE') {
                $groupPositions[] = $row;
            }
        }

        $templateGroupCount = count($groupPositions);

        // Nếu template không đủ nhóm, DUPLICATE thêm từ NHÓM 2
        if ($templateGroupCount < $groupsNeeded) {
            // Sử dụng NHÓM 2 làm mẫu (nếu có), nếu không có thì dùng nhóm 1
            $sourceGroupIndex = ($templateGroupCount > 1) ? 1 : 0;
            $sourceGroupStart = $groupPositions[$sourceGroupIndex];

            // Tính khoảng cách giữa các nhóm
            $groupHeight = ($templateGroupCount > 1)
                ? ($groupPositions[1] - $groupPositions[0])
                : 9;

            $lastGroupEnd = $groupPositions[$templateGroupCount - 1] + $groupHeight - 1;

            // Lưu thông tin ảnh từ nhóm mẫu (để clone sau)
            $sourceImages = [];
            foreach ($sheet->getDrawingCollection() as $drawing) {
                $coords = $drawing->getCoordinates();
                preg_match('/([A-Z]+)(\d+)/', $coords, $matches);

                if (isset($matches[1]) && isset($matches[2])) {
                    $col = $matches[1];
                    $row = (int)$matches[2];

                    // Nếu ảnh thuộc nhóm mẫu
                    if ($row >= $sourceGroupStart && $row < $sourceGroupStart + $groupHeight) {
                        $rowOffset = $row - $sourceGroupStart;

                        // Lưu thông tin ảnh để clone
                        $sourceImages[] = [
                            'col' => $col,
                            'rowOffset' => $rowOffset,
                            'drawing' => $drawing,
                        ];
                    }
                }
            }

            // Lưu merged cells vào array TRƯỚC để tránh lỗi "Worksheet no longer exists"
            $mergedCellsArray = [];
            foreach ($sheet->getMergeCells() as $mergeCell) {
                $mergedCellsArray[] = $mergeCell;
            }

            // Duplicate thêm các nhóm còn thiếu
            for ($groupIndex = $templateGroupCount; $groupIndex < $groupsNeeded; $groupIndex++) {
                $newGroupStart = $lastGroupEnd + 1 + (($groupIndex - $templateGroupCount) * $groupHeight);

                // Copy merged cells TRƯỚC để giữ nguyên structure
                foreach ($mergedCellsArray as $mergeCell) {
                    if (preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $mergeCell, $matches)) {
                        $startCol = $matches[1];
                        $startRow = (int)$matches[2];
                        $endCol = $matches[3];
                        $endRow = (int)$matches[4];

                        // Nếu merge cell thuộc nhóm mẫu
                        if ($startRow >= $sourceGroupStart && $startRow < $sourceGroupStart + $groupHeight) {
                            $rowOffset = $startRow - $sourceGroupStart;
                            $newStartRow = $newGroupStart + $rowOffset;
                            $newEndRow = $newStartRow + ($endRow - $startRow);

                            try {
                                $sheet->mergeCells($startCol . $newStartRow . ':' . $endCol . $newEndRow);
                            } catch (\Exception $e) {
                                // Ignore nếu đã merge hoặc conflict
                            }
                        }
                    }
                }

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

                        // Copy style bằng duplicateStyle (giữ nguyên border)
                        $sheet->duplicateStyle(
                            $sheet->getStyle($sourceCellCoord),
                            $targetCellCoord
                        );
                    }
                }

                // Copy MERGE CELLS từ source group sang target group
                $sourceMergeCells = $sheet->getMergeCells();
                foreach ($sourceMergeCells as $mergeRange) {
                    // Parse merge range (ví dụ: "A82:C82")
                    preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $mergeRange, $matches);
                    if ($matches) {
                        $startCol = $matches[1];
                        $startRow = (int)$matches[2];
                        $endCol = $matches[3];
                        $endRow = (int)$matches[4];

                        // Kiểm tra nếu merge range nằm trong source group
                        if ($startRow >= $sourceGroupStart && $endRow < ($sourceGroupStart + $groupHeight)) {
                            // Tính offset
                            $rowOffset = $startRow - $sourceGroupStart;
                            $rowEndOffset = $endRow - $sourceGroupStart;

                            // Tạo merge range mới cho target group
                            $newMergeRange = $startCol . ($newGroupStart + $rowOffset) . ':' .
                                           $endCol . ($newGroupStart + $rowEndOffset);

                            // Merge cells trong target group
                            $sheet->mergeCells($newMergeRange);
                        }
                    }
                }

                // Copy column width (chỉ cần copy 1 lần)
                if ($groupIndex == $templateGroupCount) {
                    foreach (range('A', 'I') as $col) {
                        $sheet->getColumnDimension($col)
                            ->setWidth($sheet->getColumnDimension($col)->getWidth());
                    }
                }

                // Clone ảnh/logo cho nhóm mới
                foreach ($sourceImages as $imageInfo) {
                    try {
                        $drawing = $imageInfo['drawing'];
                        $newRow = $newGroupStart + $imageInfo['rowOffset'];
                        $newCoords = $imageInfo['col'] . $newRow;

                        // Clone drawing
                        $newDrawing = clone $drawing;
                        $newDrawing->setCoordinates($newCoords);
                        $newDrawing->setWorksheet($sheet);
                    } catch (\Exception $e) {
                        // Ignore nếu không clone được
                    }
                }

                // Thêm vào danh sách vị trí nhóm
                $groupPositions[] = $newGroupStart;
            }
        }

        // Xóa dữ liệu cũ từ template gốc (chỉ xóa các nhóm template ban đầu)
        // Các nhóm duplicate đã không copy data B, E, H nên không cần xóa
        foreach ($groupPositions as $index => $groupStart) {
            // Chỉ xóa các nhóm template gốc (không phải duplicate)
            if ($index < $templateGroupCount) {
                for ($offset = 1; $offset <= 5; $offset++) {
                    foreach (['B', 'E', 'H'] as $col) {
                        $sheet->setCellValue($col . ($groupStart + $offset), null);
                    }
                }
            }
        }

        // ĐIỀN DỮ LIỆU vào từng vé (sử dụng vị trí thực tế từ template)
        foreach ($registrations as $index => $registration) {
            $groupIndex = floor($index / 3);
            $positionInGroup = $index % 3;

            // Xác định cột data
            $valueCol = ['B', 'E', 'H'][$positionInGroup];

            // Lấy vị trí thực tế của nhóm từ template
            $groupStartRow = $groupPositions[$groupIndex];

            // Điền dữ liệu (offset từ dòng "VÉ GỬI XE")
            $sheet->setCellValue($valueCol . ($groupStartRow + 1), $registration->so_ve_xe ?: '');
            $sheet->setCellValue($valueCol . ($groupStartRow + 2), $registration->ho_va_ten ?: '');
            $sheet->setCellValue($valueCol . ($groupStartRow + 3), $registration->lop ?: '');
            $sheet->setCellValue($valueCol . ($groupStartRow + 4), $registration->bien_so ?: '');
            $sheet->setCellValue($valueCol . ($groupStartRow + 5), $registration->loai_xe ?: '');
        }

        // Lưu file
        $fileName = 'Vexe_T' . $this->thang . '_' . $this->nam . '.xlsx';
        $filePath = storage_path('app/exports/' . $fileName);

        // Tạo thư mục exports nếu chưa có
        if (!file_exists(storage_path('app/exports'))) {
            mkdir(storage_path('app/exports'), 0755, true);
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);

        // TODO: Có thể notify user qua notification/event khi export xong
    }
}


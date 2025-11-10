<?php

namespace App\Exports;

use App\Models\VehicleRegistration;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\BeforeWriting;
use PhpOffice\PhpSpreadsheet\IOFactory;

class VehicleRegistrationsExport implements FromCollection, WithEvents
{
    protected $thang;
    protected $nam;

    public function __construct($thang, $nam)
    {
        $this->thang = $thang;
        $this->nam = $nam;
    }

    /**
     * Lấy dữ liệu từ database
     */
    public function collection()
    {
        return VehicleRegistration::query()
            ->where('thang', $this->thang)
            ->where('nam', $this->nam)
            ->orderBy('stt')
            ->get();
    }

    /**
     * Đăng ký events để load template và ghi dữ liệu
     */
    public function registerEvents(): array
    {
        return [
            BeforeWriting::class => function(BeforeWriting $event) {
                // Đường dẫn đến file template
                $templatePath = storage_path('Vexe_Template.xlsx');

                if (!file_exists($templatePath)) {
                    throw new \Exception('Template file not found: ' . $templatePath);
                }

                // Load template file
                $templateSpreadsheet = IOFactory::load($templatePath);
                $sheet = $templateSpreadsheet->getActiveSheet();

                // Lấy dữ liệu từ collection
                $registrations = $this->collection();

                if ($registrations->isEmpty()) {
                    $event->writer->getDelegate()->setSpreadsheet($templateSpreadsheet);
                    return;
                }

                // Tính số nhóm cần thiết (mỗi nhóm có 3 vé)
                $totalTickets = $registrations->count();
                $groupsNeeded = ceil($totalTickets / 3);

                // TÌM VỊ TRÍ các nhóm trong template (tự động phát hiện "VÉ GỬI XE")
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
                if ($templateGroupCount < $groupsNeeded && $templateGroupCount > 0) {
                    // Sử dụng nhóm cuối làm mẫu
                    $sourceGroupIndex = $templateGroupCount - 1;
                    $sourceGroupStart = $groupPositions[$sourceGroupIndex];

                    // Tính khoảng cách giữa các nhóm
                    $groupHeight = ($templateGroupCount > 1)
                        ? ($groupPositions[1] - $groupPositions[0])
                        : 9; // Mặc định 9 dòng nếu chỉ có 1 nhóm

                    $lastGroupEnd = $groupPositions[$templateGroupCount - 1] + $groupHeight - 1;

                    // Lưu merged cells để tránh lỗi iteration
                    $mergedCellsArray = [];
                    foreach ($sheet->getMergeCells() as $mergeCell) {
                        $mergedCellsArray[] = $mergeCell;
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

                            // Copy từng cell
                            foreach (range('A', 'I') as $col) {
                                $sourceCell = $sheet->getCell($col . $sourceRow);
                                $targetCell = $sheet->getCell($col . $targetRow);

                                // Copy value (chỉ label, không copy data B, E, H)
                                if (!in_array($col, ['B', 'E', 'H'])) {
                                    $targetCell->setValue($sourceCell->getValue());
                                }

                                // Copy style
                                $targetCell->getStyle()->applyFromArray(
                                    $sourceCell->getStyle()->exportArray()
                                );
                            }
                        }

                        // Copy merged cells từ nhóm mẫu
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
                                        // Ignore nếu đã merge
                                    }
                                }
                            }
                        }

                        // Thêm vào danh sách vị trí nhóm
                        $groupPositions[] = $newGroupStart;
                    }
                }

                // Xóa dữ liệu cũ từ các cột data (B, E, H)
                $highestRow = $sheet->getHighestRow();
                for ($row = 1; $row <= $highestRow; $row++) {
                    foreach (['B', 'E', 'H'] as $col) {
                        $currentValue = $sheet->getCell($col . $row)->getValue();
                        // Chỉ xóa nếu không phải label
                        if ($currentValue && !in_array($currentValue, ['Số vé xe:', 'Họ và tên:', 'Lớp:', 'Biển số:', 'Loại xe:'])) {
                            $sheet->setCellValue($col . $row, null);
                        }
                    }
                }

                // ĐIỀN DỮ LIỆU vào từng vé
                foreach ($registrations as $index => $registration) {
                    $groupIndex = floor($index / 3);
                    $positionInGroup = $index % 3;

                    // Xác định cột data (B, E, H tương ứng với vị trí 0, 1, 2)
                    $valueCol = ['B', 'E', 'H'][$positionInGroup];

                    // Lấy vị trí thực tế của nhóm
                    if (!isset($groupPositions[$groupIndex])) {
                        continue; // Skip nếu không có nhóm
                    }

                    $groupStartRow = $groupPositions[$groupIndex];

                    // Điền dữ liệu (offset từ dòng "VÉ GỬI XE")
                    $sheet->setCellValue($valueCol . ($groupStartRow + 1), $registration->so_ve_xe ?: '');
                    $sheet->setCellValue($valueCol . ($groupStartRow + 2), $registration->ho_va_ten ?: '');
                    $sheet->setCellValue($valueCol . ($groupStartRow + 3), $registration->lop ?: '');
                    $sheet->setCellValue($valueCol . ($groupStartRow + 4), $registration->bien_so ?: '');
                    $sheet->setCellValue($valueCol . ($groupStartRow + 5), $registration->loai_xe ?: '');
                }

                // Thay thế spreadsheet
                $event->writer->getDelegate()->setSpreadsheet($templateSpreadsheet);
            },
        ];
    }
}


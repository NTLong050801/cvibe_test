<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <title>Danh Sách Đăng Ký Vé Xe 2025</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: Arial, sans-serif;
            background-color: #f5f5f5;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        h1 {
            text-align: center;
            color: #333;
            margin-bottom: 30px;
            font-size: 28px;
        }

        .actions {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 15px;
        }

        .filter-form {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        .filter-form select {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
        }

        .btn-group {
            display: flex;
            gap: 10px;
        }

        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            text-decoration: none;
            display: inline-block;
            transition: background-color 0.3s;
        }

        .btn-import {
            background-color: #28a745;
            color: white;
        }

        .btn-import:hover {
            background-color: #218838;
        }

        .btn-export {
            background-color: #007bff;
            color: white;
        }

        .btn-export:hover {
            background-color: #0056b3;
        }

        .btn-filter {
            background-color: #6c757d;
            color: white;
        }

        .btn-filter:hover {
            background-color: #5a6268;
        }

        .btn-delete-all {
            background-color: #dc3545;
            color: white;
        }

        .btn-delete-all:hover {
            background-color: #c82333;
        }

        .alert {
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }

        .alert-success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }

        thead {
            background-color: #007bff;
            color: white;
        }

        th, td {
            padding: 12px;
            text-align: left;
            border: 1px solid #ddd;
        }

        th {
            font-weight: bold;
            font-size: 14px;
        }

        tbody tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        tbody tr:hover {
            background-color: #e9ecef;
        }

        .so-ve-xe {
            background-color: #ffff00;
            padding: 4px 8px;
            font-weight: bold;
            border-radius: 3px;
        }

        .pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            margin-top: 20px;
            flex-wrap: wrap;
        }

        .pagination-info {
            color: #6c757d;
            font-size: 14px;
            margin-right: 15px;
        }

        .pagination-links {
            display: flex;
            gap: 5px;
            align-items: center;
        }

        .pagination-links a,
        .pagination-links span {
            padding: 8px 12px;
            border: 1px solid #ddd;
            text-decoration: none;
            color: #007bff;
            border-radius: 4px;
            background-color: white;
            min-width: 40px;
            text-align: center;
            transition: all 0.3s;
        }

        .pagination-links .active {
            background-color: #007bff;
            color: white;
            border-color: #007bff;
            font-weight: bold;
        }

        .pagination-links .disabled {
            color: #6c757d;
            cursor: not-allowed;
            opacity: 0.5;
        }

        .pagination-links a:not(.disabled):hover {
            background-color: #e9ecef;
            border-color: #007bff;
        }

        .no-data {
            text-align: center;
            padding: 40px;
            color: #6c757d;
            font-size: 16px;
        }

        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }

        .modal-content {
            background-color: white;
            margin: 10% auto;
            padding: 30px;
            border-radius: 8px;
            width: 90%;
            max-width: 500px;
        }

        .modal-header {
            margin-bottom: 20px;
        }

        .modal-header h2 {
            color: #333;
        }

        .close {
            float: right;
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
            color: #aaa;
        }

        .close:hover {
            color: #000;
        }

        .form-group {
            margin-bottom: 15px;
        }

        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }

        .form-group input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Danh Sách Đăng Ký Vé Xe 2025</h1>

        @if(session('success'))
            <div class="alert alert-success">
                {{ session('success') }}
            </div>
        @endif

        @if(session('error'))
            <div class="alert alert-success" style="background-color: #f8d7da; color: #721c24; border-color: #f5c6cb;">
                {{ session('error') }}
            </div>
        @endif

        @if($errors->any())
            <div class="alert alert-success" style="background-color: #f8d7da; color: #721c24; border-color: #f5c6cb;">
                <ul style="margin: 0; padding-left: 20px;">
                    @foreach($errors->all() as $error)
                        <li>{{ $error }}</li>
                    @endforeach
                </ul>
            </div>
        @endif

        <div class="actions">
            <form method="GET" action="{{ route('vehicle-registrations.index') }}" class="filter-form">
                <select name="thang">
                    <option value="">-- Chọn tháng --</option>
                    @for($i = 1; $i <= 12; $i++)
                        <option value="{{ $i }}" {{ request('thang') == $i ? 'selected' : '' }}>Tháng {{ $i }}</option>
                    @endfor
                </select>
                <select name="nam">
                    <option value="">-- Chọn năm --</option>
                    @for($i = 2024; $i <= 2026; $i++)
                        <option value="{{ $i }}" {{ request('nam') == $i ? 'selected' : '' }}>{{ $i }}</option>
                    @endfor
                </select>
                <button type="submit" class="btn btn-filter">Lọc</button>
            </form>

            <div class="btn-group">
                <button onclick="openImportModal()" class="btn btn-import">📥 Import Excel</button>
                <button onclick="openExportModal()" class="btn btn-export">📤 Export Excel</button>
                <button onclick="confirmDeleteAll()" class="btn btn-delete-all">🗑️ Xóa tất cả</button>
            </div>
        </div>

        @if($registrations->count() > 0)
            <table>
                <thead>
                    <tr>
                        <th>STT</th>
                        <th>Họ và tên</th>
                        <th>Lớp</th>
                        <th>Loại xe</th>
                        <th>Biển số</th>
                        <th>Số vé xe</th>
                        <th>Ngày đăng ký</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach($registrations as $registration)
                        <tr>
                            <td>{{ $registration->stt }}</td>
                            <td>{{ $registration->ho_va_ten }}</td>
                            <td>{{ $registration->lop }}</td>
                            <td>{{ $registration->loai_xe }}</td>
                            <td>{{ $registration->bien_so }}</td>
                            <td><span class="so-ve-xe">{{ $registration->so_ve_xe }}</span></td>
                            <td>{{ $registration->ngay_dang_ky ? $registration->ngay_dang_ky->format('d/m/Y') : '' }}</td>
                        </tr>
                    @endforeach
                </tbody>
            </table>

            <div class="pagination">
                <div class="pagination-info">
                    Hiển thị {{ $registrations->firstItem() ?? 0 }} - {{ $registrations->lastItem() ?? 0 }}
                    trong tổng số {{ $registrations->total() }} kết quả
                </div>
                <div class="pagination-links">
                    {{-- Previous Page Link --}}
                    @if ($registrations->onFirstPage())
                        <span class="disabled">« Trước</span>
                    @else
                        <a href="{{ $registrations->previousPageUrl() }}">« Trước</a>
                    @endif

                    {{-- Pagination Elements --}}
                    @php
                        $start = max($registrations->currentPage() - 2, 1);
                        $end = min($start + 4, $registrations->lastPage());
                        $start = max($end - 4, 1);
                    @endphp

                    @if($start > 1)
                        <a href="{{ $registrations->url(1) }}">1</a>
                        @if($start > 2)
                            <span class="disabled">...</span>
                        @endif
                    @endif

                    @for ($i = $start; $i <= $end; $i++)
                        @if ($i == $registrations->currentPage())
                            <span class="active">{{ $i }}</span>
                        @else
                            <a href="{{ $registrations->url($i) }}">{{ $i }}</a>
                        @endif
                    @endfor

                    @if($end < $registrations->lastPage())
                        @if($end < $registrations->lastPage() - 1)
                            <span class="disabled">...</span>
                        @endif
                        <a href="{{ $registrations->url($registrations->lastPage()) }}">{{ $registrations->lastPage() }}</a>
                    @endif

                    {{-- Next Page Link --}}
                    @if ($registrations->hasMorePages())
                        <a href="{{ $registrations->nextPageUrl() }}">Tiếp »</a>
                    @else
                        <span class="disabled">Tiếp »</span>
                    @endif
                </div>
            </div>
        @else
            <div class="no-data">
                Chưa có dữ liệu đăng ký vé xe
            </div>
        @endif
    </div>

    <!-- Import Modal -->
    <div id="importModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <span class="close" onclick="closeImportModal()">&times;</span>
                <h2>Import File Excel</h2>
            </div>
            <form method="POST" action="{{ route('vehicle-registrations.import') }}" enctype="multipart/form-data">
                @csrf
                <div class="form-group">
                    <label for="file">Chọn file Excel:</label>
                    <input type="file" name="file" id="file" accept=".xlsx,.xls" required>
                </div>
                <button type="submit" class="btn btn-import" style="width: 100%;">Upload</button>
            </form>
        </div>
    </div>

    <!-- Export Modal -->
    <div id="exportModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <span class="close" onclick="closeExportModal()">&times;</span>
                <h2>Export File Excel</h2>
            </div>
            <form method="POST" action="{{ route('vehicle-registrations.export') }}" id="exportForm">
                @csrf
                <div class="form-group">
                    <label for="export_nam">Chọn năm:</label>
                    <select name="nam" id="export_nam" required style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">
                        <option value="">-- Chọn năm --</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="export_thang">Chọn tháng:</label>
                    <select name="thang" id="export_thang" required style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">
                        <option value="">-- Chọn tháng trước --</option>
                    </select>
                </div>
                <button type="submit" class="btn btn-export" style="width: 100%;">Export</button>
            </form>
        </div>
    </div>

    <script>
        let availableData = {};

        // Load available months khi trang load
        document.addEventListener('DOMContentLoaded', function() {
            loadAvailableMonths();
        });

        function loadAvailableMonths() {
            fetch('{{ route("vehicle-registrations.available-months") }}')
                .then(response => response.json())
                .then(data => {
                    availableData = data;
                })
                .catch(error => {
                    console.error('Error loading available months:', error);
                });
        }

        function openImportModal() {
            document.getElementById('importModal').style.display = 'block';
        }

        function closeImportModal() {
            document.getElementById('importModal').style.display = 'none';
        }

        function openExportModal() {
            const modal = document.getElementById('exportModal');
            const namSelect = document.getElementById('export_nam');
            const thangSelect = document.getElementById('export_thang');

            // Reset form
            namSelect.innerHTML = '<option value="">-- Chọn năm --</option>';
            thangSelect.innerHTML = '<option value="">-- Chọn tháng trước --</option>';
            thangSelect.disabled = true;

            // Populate năm dropdown
            const years = Object.keys(availableData).sort((a, b) => b - a);
            if (years.length === 0) {
                alert('Không có dữ liệu để export!');
                return;
            }

            years.forEach(year => {
                const option = document.createElement('option');
                option.value = year;
                option.textContent = year;
                namSelect.appendChild(option);
            });

            modal.style.display = 'block';
        }

        function closeExportModal() {
            document.getElementById('exportModal').style.display = 'none';
        }

        // Khi chọn năm, load các tháng tương ứng
        document.getElementById('export_nam').addEventListener('change', function() {
            const thangSelect = document.getElementById('export_thang');
            const selectedYear = this.value;

            thangSelect.innerHTML = '<option value="">-- Chọn tháng --</option>';

            if (selectedYear && availableData[selectedYear]) {
                thangSelect.disabled = false;
                const months = availableData[selectedYear].sort((a, b) => b - a);

                months.forEach(month => {
                    const option = document.createElement('option');
                    option.value = month;
                    option.textContent = 'Tháng ' + month;
                    thangSelect.appendChild(option);
                });
            } else {
                thangSelect.disabled = true;
            }
        });

        function confirmDeleteAll() {
            if (confirm('⚠️ BẠN CÓ CHẮC CHẮN MUỐN XÓA TẤT CẢ DỮ LIỆU?\n\nHành động này không thể hoàn tác!')) {
                // Tạo form để submit
                const form = document.createElement('form');
                form.method = 'POST';
                form.action = '{{ route("vehicle-registrations.delete-all") }}';

                // Thêm CSRF token
                const csrfInput = document.createElement('input');
                csrfInput.type = 'hidden';
                csrfInput.name = '_token';
                csrfInput.value = '{{ csrf_token() }}';
                form.appendChild(csrfInput);

                // Thêm method DELETE
                const methodInput = document.createElement('input');
                methodInput.type = 'hidden';
                methodInput.name = '_method';
                methodInput.value = 'DELETE';
                form.appendChild(methodInput);

                document.body.appendChild(form);
                form.submit();
            }
        }

        window.onclick = function(event) {
            const importModal = document.getElementById('importModal');
            const exportModal = document.getElementById('exportModal');

            if (event.target == importModal) {
                importModal.style.display = 'none';
            }
            if (event.target == exportModal) {
                exportModal.style.display = 'none';
            }
        }
    </script>
</body>
</html>


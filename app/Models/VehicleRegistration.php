<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class VehicleRegistration extends Model
{
    protected $fillable = [
        'stt',
        'ho_va_ten',
        'lop',
        'loai_xe',
        'bien_so',
        'so_ve_xe',
        'ngay_dang_ky',
        'thang',
        'nam',
    ];

    protected $casts = [
        'ngay_dang_ky' => 'date',
    ];
}

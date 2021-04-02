<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomerInformation extends Model
{
    use SoftDeletes;

    public  $table = 'customer_information';
}

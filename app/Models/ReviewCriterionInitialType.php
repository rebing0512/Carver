<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class ReviewCriterionInitialType extends Model
{
    use SoftDeletes;

    public $table = 'review_criterion_initial_type';

    public function review(){
        return $this->hasMany(ReviewCriterionInitial::class,'type_id','id');
    }
}

<?php


namespace App\Repositories;


use App\Models\Homepage\Users;

class EmailRepo implements IEmailRepo
{
    public function getUser(){
        return Users::first();
    }

}
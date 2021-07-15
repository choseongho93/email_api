<?php

namespace App\Providers;

use Illuminate\Support\ServiceProvider;
use App\Repositories\IEmailRepo;
use App\Repositories\EmailRepo;

class RepositoryServiceProvider extends ServiceProvider
{
    /**
     * Register services.
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind(IEmailRepo::class,EmailRepo::class);
    }

    /**
     * Bootstrap services.
     *
     * @return void
     */
    public function boot()
    {
        //
    }
}

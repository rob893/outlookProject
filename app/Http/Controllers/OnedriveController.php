<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;

class OnedriveController extends Controller
{
    public function addRow(){
        if (session_status() == PHP_SESSION_NONE) {
            session_start();
        }
        
        $tokenCache = new \App\TokenStore\TokenCache;
        
        $graph = new Graph();
        $graph->setAccessToken($tokenCache->getAccessToken());
        
        $user = $graph->createRequest('GET', '/me')
            ->setReturnType(Model\User::class)
            ->execute();
        
        
        $messageQueryParams =  '{
            "index":  null,
            "values": [
                ["alex darrow", "adarrow@tenant.onmicrosoft.com"]
            ]
        }';
        
        $getItemsUrl = '/me/drive/root:/demo.xlsx';
        $addRowUrl = '/me/drive/root:/demo.xlsx:/workbook/tables/Table1/rows/add';
        $messages = $graph->createRequest('POST', $addRowUrl)
            ->attachBody($messageQueryParams)
            ->execute();
        
        echo '<pre>';
        print_r($messages->getBody());
        echo '</pre>';
    }
}
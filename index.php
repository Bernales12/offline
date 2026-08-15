<?php
session_start();
date_default_timezone_set('Asia/Manila');

/*
|--------------------------------------------------------------------------
| PHARMACY MEDICINE INVENTORY SYSTEM - NO DATABASE
|--------------------------------------------------------------------------
| Single-file PHP application.
| Data is stored in PHP session only.
| No MySQL / MariaDB / PDO / config.php required.
|--------------------------------------------------------------------------
*/

$DEFAULT_LOW_STOCK = 200;
$EXPIRY_WARNING_DAYS = 30;

if (!isset($_SESSION['medicines'])) {
    $_SESSION['medicines'] = [
        'M001' => [
            'sku'=>'M001','inventory_name'=>'Paracetamol','strength'=>'500','unit'=>'mg',
            'dosage_form'=>'Tablet','generic_name'=>'Paracetamol (Acetaminophen)',
            'quantity'=>150,'batch_number'=>'BCH-2026-01A','expiration_date'=>'2028-05-15',
            'category'=>'Analgesics','low_stock_threshold'=>200
        ],
        'M002' => [
            'sku'=>'M002','inventory_name'=>'Amoxicillin','strength'=>'500','unit'=>'mg',
            'dosage_form'=>'Capsule','generic_name'=>'Amoxicillin Trihydrate',
            'quantity'=>12,'batch_number'=>'BCH-2025-09C','expiration_date'=>'2027-11-20',
            'category'=>'Antibiotics','low_stock_threshold'=>20
        ],
        'M003' => [
            'sku'=>'M003','inventory_name'=>'Cetirizine','strength'=>'10','unit'=>'mg',
            'dosage_form'=>'Tablet','generic_name'=>'Cetirizine Dihydrochloride',
            'quantity'=>8,'batch_number'=>'BCH-2026-03X','expiration_date'=>'2028-01-10',
            'category'=>'Antihistamines','low_stock_threshold'=>15
        ]
    ];
}

if (!isset($_SESSION['dispense_logs'])) {
    $_SESSION['dispense_logs'] = [
        [
            'date'=>'20 May 2026','date_iso'=>'2026-05-20',
            'inventory_name'=>'Paracetamol 500 mg Tablet',
            'batch_number'=>'BCH-2026-01A','qty_out'=>10,'recipient'=>'John Doe'
        ]
    ];
}

function h($v) {
    return htmlspecialchars((string)$v, ENT_QUOTES, 'UTF-8');
}

function medicineFullName($m) {
    return trim(($m['inventory_name']??'').' '.($m['strength']??'').' '.($m['unit']??'').' '.($m['dosage_form']??''));
}

function getThreshold($m) {
    $n = (int)($m['low_stock_threshold'] ?? 200);
    return $n < 1 ? 200 : $n;
}

function generateMedicineKey() {
    do {
        $key = 'M'.random_int(1000,9999);
    } while (isset($_SESSION['medicines'][$key]));
    return $key;
}

function sortMedicines() {
    uasort($_SESSION['medicines'], function($a,$b) {
        return strcasecmp($a['inventory_name']??'', $b['inventory_name']??'');
    });
}

function processExpiredMedicines() {
    $today = date('Y-m-d');
    foreach ($_SESSION['medicines'] as &$m) {
        if (!empty($m['expiration_date']) && $m['expiration_date'] <= $today) {
            $m['quantity'] = 0;
        }
    }
    unset($m);
}

function saveInventoryExcel($inventory) {
    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename=medicine_inventory.xls');
    echo '<html><head><meta charset="UTF-8"></head><body>';
    echo '<h2>Medicine Inventory Report</h2>';
    echo '<p>Generated: '.h(date('d M Y H:i:s')).'</p>';
    echo '<table border="1"><tr>
    <th>SKU</th><th>Medicine</th><th>Strength</th><th>Unit</th><th>Dosage Form</th>
    <th>Generic Name</th><th>Category</th><th>Batch</th><th>Expiration</th>
    <th>Quantity</th><th>Low Stock Threshold</th><th>Status</th></tr>';

    $today = strtotime(date('Y-m-d'));
    $warning = strtotime('+30 days',$today);

    foreach ($inventory as $m) {
        $qty=(int)$m['quantity']; $threshold=getThreshold($m);
        $exp=!empty($m['expiration_date'])?strtotime($m['expiration_date']):false;
        if ($exp!==false && $exp<=$today) $status='EXPIRED';
        elseif ($exp!==false && $exp<=$warning) $status='EXPIRING SOON';
        elseif ($qty<=$threshold) $status='LOW STOCK';
        else $status='AVAILABLE';

        echo '<tr>';
        foreach([
            $m['sku']??'', $m['inventory_name']??'', $m['strength']??'', $m['unit']??'',
            $m['dosage_form']??'', $m['generic_name']??'', $m['category']??'',
            $m['batch_number']??'', $m['expiration_date']??'', $qty,$threshold,$status
        ] as $v) echo '<td>'.h($v).'</td>';
        echo '</tr>';
    }
    echo '</table></body></html>';
    exit;
}

function saveDispensingExcel($month,$year,$logs) {
    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename=medicine_dispensing_'.$year.'_'.str_pad($month,2,'0',STR_PAD_LEFT).'.xls');
    echo '<html><head><meta charset="UTF-8"></head><body>';
    echo '<h2>Medicine Dispensing Report</h2>';
    echo '<h3>'.h(date('F Y',strtotime("$year-$month-01"))).'</h3>';
    echo '<table border="1"><tr><th>Date</th><th>Medicine</th><th>Batch</th><th>Quantity Dispensed</th><th>Recipient</th></tr>';

    foreach($logs as $l) {
        $t=strtotime($l['date_iso']??'');
        if (!$t || (int)date('m',$t)!=$month || (int)date('Y',$t)!=$year) continue;
        echo '<tr><td>'.h($l['date']??'').'</td><td>'.h($l['inventory_name']??'').'</td><td>'.
             h($l['batch_number']??'').'</td><td>'.(int)($l['qty_out']??0).'</td><td>'.h($l['recipient']??'').'</td></tr>';
    }
    echo '</table></body></html>';
    exit;
}

/* POST ACTIONS */
$message='';
$messageType='success';

if ($_SERVER['REQUEST_METHOD']==='POST') {
    $action=$_POST['action']??'';

    if ($action==='reset_demo') {
        session_unset();
        header('Location: '.$_SERVER['PHP_SELF']);
        exit;
    }

    if ($action==='add_medicine') {
        $name=trim($_POST['inventory_name']??'');
        if ($name==='') {
            $message='Medicine name is required.';
            $messageType='danger';
        } else {
            $key=generateMedicineKey();
            $_SESSION['medicines'][$key]=[
                'sku'=>$key,
                'inventory_name'=>$name,
                'strength'=>trim($_POST['strength']??''),
                'unit'=>trim($_POST['unit']??'mg'),
                'dosage_form'=>trim($_POST['dosage_form']??''),
                'generic_name'=>trim($_POST['generic_name']??''),
                'quantity'=>max(0,(int)($_POST['quantity']??0)),
                'batch_number'=>trim($_POST['batch_number']??''),
                'expiration_date'=>$_POST['expiration_date']??'',
                'category'=>trim($_POST['category']??'General'),
                'low_stock_threshold'=>max(1,(int)($_POST['low_stock_threshold']??$DEFAULT_LOW_STOCK))
            ];
            $message='Medicine successfully added.';
        }
    }

    elseif ($action==='edit_medicine') {
        $key=$_POST['medicine_key']??'';
        if (isset($_SESSION['medicines'][$key])) {
            $_SESSION['medicines'][$key]['inventory_name']=trim($_POST['inventory_name']??'');
            $_SESSION['medicines'][$key]['strength']=trim($_POST['strength']??'');
            $_SESSION['medicines'][$key]['unit']=trim($_POST['unit']??'mg');
            $_SESSION['medicines'][$key]['dosage_form']=trim($_POST['dosage_form']??'');
            $_SESSION['medicines'][$key]['generic_name']=trim($_POST['generic_name']??'');
            $_SESSION['medicines'][$key]['quantity']=max(0,(int)($_POST['quantity']??0));
            $_SESSION['medicines'][$key]['batch_number']=trim($_POST['batch_number']??'');
            $_SESSION['medicines'][$key]['expiration_date']=$_POST['expiration_date']??'';
            $_SESSION['medicines'][$key]['category']=trim($_POST['category']??'General');
            $_SESSION['medicines'][$key]['low_stock_threshold']=max(1,(int)($_POST['low_stock_threshold']??$DEFAULT_LOW_STOCK));
            $message='Medicine information updated.';
        } else {
            $message='Medicine could not be found.';
            $messageType='danger';
        }
    }

    elseif ($action==='delete_medicine') {
        $key=$_POST['medicine_key']??'';
        if (isset($_SESSION['medicines'][$key])) {
            $name=$_SESSION['medicines'][$key]['inventory_name'];
            unset($_SESSION['medicines'][$key]);
            $message=$name.' was removed from inventory.';
        } else {
            $message='Medicine could not be found.';
            $messageType='danger';
        }
    }

    elseif ($action==='stock_out') {
        processExpiredMedicines();
        $key=$_POST['medicine_key']??'';
        $qty=max(0,(int)($_POST['qty_out']??0));
        $recipient=trim($_POST['recipient']??'');

        if (!isset($_SESSION['medicines'][$key])) {
            $message='Please select a valid medicine.'; $messageType='danger';
        } elseif ($qty<1) {
            $message='Quantity must be at least 1.'; $messageType='danger';
        } elseif ($recipient==='') {
            $message='Patient / Recipient is required.'; $messageType='danger';
        } else {
            $m=&$_SESSION['medicines'][$key];
            $expired=!empty($m['expiration_date']) && $m['expiration_date']<=date('Y-m-d');

            if ($expired) {
                $message='This medicine has expired and cannot be dispensed.'; $messageType='danger';
            } elseif ((int)$m['quantity']<$qty) {
                $message='Error: Stock is not enough for this medicine.'; $messageType='danger';
            } else {
                $m['quantity']-=$qty;
                $full=medicineFullName($m);
                $_SESSION['dispense_logs'][]=[
                    'date'=>date('d M Y'),'date_iso'=>date('Y-m-d'),
                    'inventory_name'=>$full,'batch_number'=>$m['batch_number']??'',
                    'qty_out'=>$qty,'recipient'=>$recipient
                ];
                $message='Successfully dispensed '.$qty.' unit(s) of '.$full.'.';
            }
            unset($m);
        }
    }
}

processExpiredMedicines();
sortMedicines();

if (isset($_GET['export']) && $_GET['export']==='inventory') {
    saveInventoryExcel($_SESSION['medicines']);
}
if (isset($_GET['export']) && $_GET['export']==='dispensing') {
    saveDispensingExcel((int)($_GET['month']??date('m')),(int)($_GET['year']??date('Y')),array_reverse($_SESSION['dispense_logs']));
}

$medicineInventory=$_SESSION['medicines'];
$dispenseLogs=array_reverse($_SESSION['dispense_logs']);

$today=strtotime(date('Y-m-d'));
$expiryThreshold=strtotime("+{$EXPIRY_WARNING_DAYS} days",$today);

$totalProducts=count($medicineInventory);
$totalStock=0;
$lowStockCount=0;
$expiredCount=0;
$expiringCount=0;
$categoryCounts=[];
$expiredMedicines=[];
$expiringMedicines=[];
$lowStockMedicines=[];

foreach($medicineInventory as $key=>$m) {
    $qty=(int)$m['quantity'];
    $totalStock+=$qty;
    $threshold=getThreshold($m);
    $exp=!empty($m['expiration_date'])?strtotime($m['expiration_date']):false;

    if($exp!==false && $exp<=$today) {
        $expiredCount++; $expiredMedicines[$key]=$m;
    } elseif($exp!==false && $exp<=$expiryThreshold) {
        $expiringCount++; $expiringMedicines[$key]=$m;
    }

    if($qty<=$threshold && !($exp!==false && $exp<=$today)) {
        $lowStockCount++; $lowStockMedicines[$key]=$m;
    }

    $cat=trim($m['category']??'')?:'General';
    $categoryCounts[$cat]=($categoryCounts[$cat]??0)+$qty;
}

$selectedMonth=(int)($_GET['report_month']??date('m'));
$selectedYear=(int)($_GET['report_year']??date('Y'));
$monthlyDispense=[];
$monthlyTotal=0;
$monthlyMedicineTotals=[];

foreach($dispenseLogs as $l) {
    $t=strtotime($l['date_iso']??'');
    if($t && (int)date('m',$t)===$selectedMonth && (int)date('Y',$t)===$selectedYear) {
        $monthlyDispense[]=$l;
        $q=(int)($l['qty_out']??0);
        $monthlyTotal+=$q;
        $n=$l['inventory_name']??'Unknown';
        $monthlyMedicineTotals[$n]=($monthlyMedicineTotals[$n]??0)+$q;
    }
}
arsort($monthlyMedicineTotals);

$trendLabels=[];$trendData=[];
for($i=5;$i>=0;$i--) {
    $ts=strtotime("-$i months",$today);
    $m=(int)date('m',$ts);$y=(int)date('Y',$ts);
    $trendLabels[]=date('M Y',$ts);$sum=0;
    foreach($dispenseLogs as $l) {
        $t=strtotime($l['date_iso']??'');
        if($t && (int)date('m',$t)===$m && (int)date('Y',$t)===$y) $sum+=(int)$l['qty_out'];
    }
    $trendData[]=$sum;
}

$availableCount=max(0,$totalProducts-$lowStockCount-$expiringCount-$expiredCount);
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharmacy Inventory - No Database</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
:root{--p950:#150a2b;--p900:#1f1040;--p800:#2b1656;--p600:#5b21b6;--p500:#7c3aed;--p300:#b18ffb;--p100:#f0eafd;--bg:#f5f3fb;--card:#fff;--border:#eae4f7;--text:#241a3d;--muted:#8b81a3;--success:#16a34a;--warning:#d97706;--danger:#e11d48}
*{font-family:"Plus Jakarta Sans",system-ui,sans-serif}body{background:var(--bg);color:var(--text)}
.sidebar{background:linear-gradient(190deg,var(--p950),var(--p900) 45%,var(--p800));min-height:100vh;color:#d9c8fb;padding:20px 16px}
.logo{display:flex;gap:12px;align-items:center;padding:6px 8px 22px}.logoicon{width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,#9061f9,#ec4899);display:flex;align-items:center;justify-content:center;color:white}.logotext strong{display:block;color:white;font-weight:800}.logotext small{font-size:10px;color:#b18ffb}
.nav-link{color:#d9c8fb!important;border-radius:10px;margin-bottom:4px;padding:11px 16px;font-size:14px}.nav-link:hover{background:#ffffff10}.nav-link.active{background:linear-gradient(135deg,#7c3aed,#5b21b6)!important;color:white!important;box-shadow:0 6px 16px #7c3aed66}.insights{background:#ffffff0d;border:1px solid #ffffff14;border-radius:12px;padding:14px;margin-top:14px;font-size:12px}.insights li{margin:7px 0}.footer{font-size:11px;border-top:1px solid #ffffff14;padding-top:12px}
.main{min-height:100vh}.top{background:#fff;border-bottom:1px solid var(--border)}.cardx{background:#fff;border:1px solid var(--border);border-radius:16px;box-shadow:0 1px 2px #4c1d950a}.kpi{display:flex;gap:14px;padding:20px}.ki{width:46px;height:46px;min-width:46px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;background:linear-gradient(135deg,#9061f9,#5b21b6)}.ki.warn{background:linear-gradient(135deg,#fbbf24,#d97706)}.ki.danger{background:linear-gradient(135deg,#fb7185,#e11d48)}.ki.dark{background:linear-gradient(135deg,#2b1656,#150a2b)}.kbody small{color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase}.num{font-size:26px;font-weight:800}.table th{background:var(--p100);font-size:11px;text-transform:uppercase;color:#5b4b75}.table td{vertical-align:middle;font-size:13px}.form-control,.form-select{border-color:#ddd2f7;border-radius:9px}.form-control:focus,.form-select:focus{border-color:#9061f9;box-shadow:0 0 .2rem #7c3aed29}.btn-primary{background:linear-gradient(135deg,#7c3aed,#5b21b6);border:0}.expired{background:#fff1f2!important}.expiring{background:#fff7ed!important}.low{background:#fffbeb!important}.modal-header{background:linear-gradient(135deg,#1f1040,#5b21b6);color:#fff}.modal-title{color:#fff}.alert{border-radius:12px}
@media(max-width:768px){.sidebar{min-height:auto}.num{font-size:21px}}
</style>
</head>
<body>
<div class="container-fluid p-0">
<div class="row g-0">
<aside class="col-md-2 sidebar d-flex flex-column justify-content-between">
<div>
<div class="logo"><div class="logoicon"><i class="fa-solid fa-pills"></i></div><div class="logotext"><strong>PHARMACY</strong><small>NO DATABASE</small></div></div>
<ul class="nav nav-pills flex-column">
<li><button class="nav-link active w-100 text-start" data-bs-toggle="pill" data-bs-target="#dashboard"><i class="fa-solid fa-gauge-high me-2"></i>Dashboard</button></li>
<li><button class="nav-link w-100 text-start" data-bs-toggle="pill" data-bs-target="#inventory"><i class="fa-solid fa-boxes-stacked me-2"></i>Medicine Inventory</button></li>
<li><button class="nav-link w-100 text-start" data-bs-toggle="pill" data-bs-target="#stockout"><i class="fa-solid fa-truck-ramp-box me-2"></i>Dispense / Stock-Out</button></li>
<li><button class="nav-link w-100 text-start" data-bs-toggle="pill" data-bs-target="#reports"><i class="fa-solid fa-chart-pie me-2"></i>Reports & Analytics</button></li>
</ul>
<div class="insights"><strong class="text-white">QUICK INSIGHTS</strong><ul class="list-unstyled mt-2 mb-0">
<li>Total stock: <strong><?=number_format($totalStock)?></strong> units</li>
<li><?=$lowStockCount?> item(s) low stock</li>
<li><?=$expiredCount?> expired</li>
<li><?=$expiringCount?> expiring within 30 days</li>
</ul></div>
</div>
<div class="footer">Data as of <?=date('M j, Y g:i A')?></div>
</aside>

<main class="col-md-10 main">
<div class="top px-4 py-3 d-flex justify-content-between align-items-center">
<div><h4 class="mb-0 fw-bold">Pharmacy Dashboard</h4><small class="text-muted">No MySQL / No database required</small></div>
<div>
<a href="?export=inventory" class="btn btn-success btn-sm me-2"><i class="fa-solid fa-file-excel me-1"></i>Inventory Excel</a>
<a href="?export=dispensing&month=<?=date('m')?>&year=<?=date('Y')?>" class="btn btn-danger btn-sm"><i class="fa-solid fa-file-excel me-1"></i>Dispensing Excel</a>
</div>
</div>

<div class="p-4">
<?php if($message): ?><div class="alert alert-<?=$messageType?> alert-dismissible fade show"><?=h($message)?><button class="btn-close" data-bs-dismiss="alert"></button></div><?php endif; ?>

<div class="tab-content">
<section class="tab-pane fade show active" id="dashboard">
<div class="row g-3 mb-4">
<?php
$cards=[
['fa-capsules','Total Medicines',$totalProducts,'Active items',''],
['fa-layer-group','Total Stock On Hand',number_format($totalStock),'Units in inventory','dark'],
['fa-triangle-exclamation','Low Stock',$lowStockCount,'Needs reorder','warn'],
['fa-calendar-xmark','Expired',$expiredCount,'Stock zeroed','danger']
];
foreach($cards as $c): ?>
<div class="col-xl-3 col-md-6"><div class="cardx kpi"><div class="ki <?=$c[4]?>"><i class="fa-solid <?=$c[0]?>"></i></div><div class="kbody"><small><?=$c[1]?></small><div class="num"><?=$c[2]?></div><span class="badge bg-light text-secondary"><?=$c[3]?></span></div></div></div>
<?php endforeach; ?>
</div>

<div class="row g-3 mb-4">
<div class="col-lg-5"><div class="cardx p-4"><h6 class="fw-bold">Dispensing Trend (6 Months)</h6><canvas id="trendChart"></canvas></div></div>
<div class="col-lg-4"><div class="cardx p-4"><h6 class="fw-bold">Stock by Category</h6><canvas id="categoryChart"></canvas></div></div>
<div class="col-lg-3"><div class="cardx p-4"><h6 class="fw-bold">Stock Status</h6><canvas id="statusChart"></canvas></div></div>
</div>

<div class="row g-3 mb-4">
<div class="col-lg-6"><div class="cardx p-4"><h5 class="fw-bold text-danger">Expired Medicines <span class="badge bg-danger"><?=$expiredCount?></span></h5>
<?php if(!$expiredMedicines): ?><p class="text-muted py-3">No expired medicines.</p><?php else: ?><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Medicine</th><th>Expiration</th><th>Stock</th></tr></thead><tbody>
<?php foreach($expiredMedicines as $m): ?><tr class="expired"><td class="fw-bold"><?=h(medicineFullName($m))?></td><td><?=h($m['expiration_date'])?></td><td>0</td></tr><?php endforeach; ?>
</tbody></table></div><?php endif; ?></div></div>

<div class="col-lg-6"><div class="cardx p-4"><h5 class="fw-bold text-warning">Expiring Within 30 Days <span class="badge bg-warning text-dark"><?=$expiringCount?></span></h5>
<?php if(!$expiringMedicines): ?><p class="text-muted py-3">No medicine will expire within 30 days.</p><?php else: ?><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Medicine</th><th>Expiration</th><th>Days Left</th></tr></thead><tbody>
<?php foreach($expiringMedicines as $m): $days=ceil((strtotime($m['expiration_date'])-$today)/86400); ?><tr class="expiring"><td class="fw-bold"><?=h(medicineFullName($m))?></td><td><?=h($m['expiration_date'])?></td><td><span class="badge bg-warning text-dark"><?=$days?> day(s)</span></td></tr><?php endforeach; ?>
</tbody></table></div><?php endif; ?></div></div>
</div>

<div class="row g-3">
<div class="col-lg-6"><div class="cardx p-4"><h5 class="fw-bold">Low Stock Alerts</h5>
<?php if(!$lowStockMedicines): ?><p class="text-muted">No low-stock medicines.</p><?php else: foreach($lowStockMedicines as $m): ?><div class="d-flex justify-content-between border-bottom py-2"><div><strong><?=h(medicineFullName($m))?></strong><br><small class="text-muted"><?=h($m['category'])?> · Batch <?=h($m['batch_number'])?></small></div><div class="text-end"><strong class="text-danger"><?=$m['quantity']?> left</strong><br><small>Reorder at <?=getThreshold($m)?></small></div></div><?php endforeach; endif; ?>
</div></div>
<div class="col-lg-6"><div class="cardx p-4"><h5 class="fw-bold">Recent Medicines Dispensed</h5><div class="table-responsive"><table class="table"><thead><tr><th>Date</th><th>Medicine</th><th>Qty</th><th>Recipient</th></tr></thead><tbody>
<?php foreach(array_slice($dispenseLogs,0,6) as $l): ?><tr><td><?=h($l['date'])?></td><td><?=h($l['inventory_name'])?></td><td class="text-danger fw-bold">-<?=$l['qty_out']?></td><td><?=h($l['recipient'])?></td></tr><?php endforeach; ?>
</tbody></table></div></div></div>
</div>
</section>

<section class="tab-pane fade" id="inventory">
<div class="d-flex justify-content-between mb-3"><div><h4 class="fw-bold">Medicine Inventory</h4><small class="text-muted">All inventory is stored in the PHP session.</small></div><a href="?export=inventory" class="btn btn-success"><i class="fa-solid fa-file-excel me-1"></i>Inventory Excel</a></div>
<div class="cardx p-4 mb-4"><h5 class="fw-bold">Register New Medicine</h5>
<form method="post"><input type="hidden" name="action" value="add_medicine"><div class="row g-3">
<div class="col-md-3"><label class="form-label">Medicine Name</label><input name="inventory_name" class="form-control" required></div>
<div class="col-md-2"><label class="form-label">Strength</label><input name="strength" class="form-control" required></div>
<div class="col-md-2"><label class="form-label">Unit</label><select name="unit" class="form-select"><option>mg</option><option>ml</option><option>g</option><option>mcg</option><option>%</option></select></div>
<div class="col-md-2"><label class="form-label">Dosage Form</label><input name="dosage_form" class="form-control" placeholder="Tablet" required></div>
<div class="col-md-3"><label class="form-label">Generic Name</label><input name="generic_name" class="form-control" required></div>
<div class="col-md-3"><label class="form-label">Category</label><input name="category" class="form-control" value="General" required></div>
<div class="col-md-2"><label class="form-label">Quantity</label><input type="number" name="quantity" class="form-control" min="0" value="100" required></div>
<div class="col-md-2"><label class="form-label">Batch Number</label><input name="batch_number" class="form-control" required></div>
<div class="col-md-3"><label class="form-label">Expiration Date</label><input type="date" name="expiration_date" class="form-control" required></div>
<div class="col-md-2"><label class="form-label">Low Stock At</label><input type="number" name="low_stock_threshold" class="form-control" min="1" value="200" required></div>
<div class="col-12"><button class="btn btn-primary"><i class="fa-solid fa-plus me-1"></i>Add Medicine</button></div>
</div></form></div>

<div class="cardx p-4"><h5 class="fw-bold">Complete Medicine List <span class="badge bg-primary"><?=$totalProducts?></span></h5>
<div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Medicine</th><th>Strength</th><th>Form</th><th>Category</th><th>Batch</th><th>Expiration</th><th>Stock</th><th>Low At</th><th>Status</th><th>Actions</th></tr></thead><tbody>
<?php foreach($medicineInventory as $key=>$m):
$qty=(int)$m['quantity'];$th=getThreshold($m);$exp=!empty($m['expiration_date'])?strtotime($m['expiration_date']):false;
if($exp!==false&&$exp<=$today){$status='<span class="badge bg-dark">EXPIRED</span>';$rc='expired';}
elseif($exp!==false&&$exp<=$expiryThreshold){$status='<span class="badge bg-warning text-dark">EXPIRING SOON</span>';$rc='expiring';}
elseif($qty<=$th){$status='<span class="badge bg-warning text-dark">LOW STOCK</span>';$rc='low';}
else{$status='<span class="badge bg-success">AVAILABLE</span>';$rc='';}
?>
<tr class="<?=$rc?>"><td class="fw-bold"><?=h(medicineFullName($m))?><br><small class="text-muted">SKU: <?=h($key)?></small></td><td><?=h($m['strength'])?> <?=h($m['unit'])?></td><td><?=h($m['dosage_form'])?></td><td><?=h($m['category'])?></td><td><?=h($m['batch_number'])?></td><td><?=h($m['expiration_date'])?></td><td class="fw-bold"><?=$qty?></td><td><?=$th?></td><td><?=$status?></td><td>
<button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#edit<?=h($key)?>"><i class="fa-solid fa-pen"></i></button>
<form method="post" class="d-inline" onsubmit="return confirm('Remove this medicine?')"><input type="hidden" name="action" value="delete_medicine"><input type="hidden" name="medicine_key" value="<?=h($key)?>"><button class="btn btn-sm btn-danger"><i class="fa-solid fa-trash"></i></button></form>
</td></tr>
<?php endforeach; ?></tbody></table></div></div>

<?php foreach($medicineInventory as $key=>$m): ?>
<div class="modal fade" id="edit<?=h($key)?>" tabindex="-1"><div class="modal-dialog modal-lg"><div class="modal-content"><div class="modal-header"><h5 class="modal-title">Edit Medicine</h5><button class="btn-close" data-bs-dismiss="modal"></button></div>
<form method="post"><input type="hidden" name="action" value="edit_medicine"><input type="hidden" name="medicine_key" value="<?=h($key)?>"><div class="modal-body"><div class="row g-3">
<div class="col-md-4"><label class="form-label">Medicine Name</label><input name="inventory_name" class="form-control" value="<?=h($m['inventory_name'])?>" required></div>
<div class="col-md-2"><label class="form-label">Strength</label><input name="strength" class="form-control" value="<?=h($m['strength'])?>" required></div>
<div class="col-md-2"><label class="form-label">Unit</label><select name="unit" class="form-select"><?php foreach(['mg','ml','g','mcg','%'] as $u): ?><option <?=$m['unit']===$u?'selected':''?>><?=$u?></option><?php endforeach; ?></select></div>
<div class="col-md-4"><label class="form-label">Dosage Form</label><input name="dosage_form" class="form-control" value="<?=h($m['dosage_form'])?>" required></div>
<div class="col-md-6"><label class="form-label">Generic Name</label><input name="generic_name" class="form-control" value="<?=h($m['generic_name'])?>" required></div>
<div class="col-md-6"><label class="form-label">Category</label><input name="category" class="form-control" value="<?=h($m['category'])?>" required></div>
<div class="col-md-3"><label class="form-label">Quantity</label><input type="number" name="quantity" class="form-control" min="0" value="<?=$m['quantity']?>" required></div>
<div class="col-md-3"><label class="form-label">Batch</label><input name="batch_number" class="form-control" value="<?=h($m['batch_number'])?>" required></div>
<div class="col-md-3"><label class="form-label">Expiration</label><input type="date" name="expiration_date" class="form-control" value="<?=h($m['expiration_date'])?>" required></div>
<div class="col-md-3"><label class="form-label">Low Stock</label><input type="number" name="low_stock_threshold" class="form-control" min="1" value="<?=getThreshold($m)?>" required></div>
</div></div><div class="modal-footer"><button class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button><button class="btn btn-primary">Save Changes</button></div></form>
</div></div></div>
<?php endforeach; ?>
</section>

<section class="tab-pane fade" id="stockout">
<div class="d-flex justify-content-between mb-3"><div><h4 class="fw-bold">Dispense / Stock-Out</h4><small class="text-muted">Transactions are stored in the current PHP session.</small></div><a href="?export=dispensing&month=<?=date('m')?>&year=<?=date('Y')?>" class="btn btn-danger">Current Month Excel</a></div>
<div class="cardx p-4 mb-4"><h5 class="fw-bold">Dispense Medicine</h5><form method="post"><input type="hidden" name="action" value="stock_out"><div class="row g-3">
<div class="col-md-6"><label class="form-label">Select Medicine</label><select name="medicine_key" class="form-select" required><option value="">Choose medicine...</option><?php foreach($medicineInventory as $key=>$m): $ex=!empty($m['expiration_date'])&&$m['expiration_date']<=date('Y-m-d'); ?><option value="<?=h($key)?>" <?=$ex?'disabled':''?>><?=h(medicineFullName($m).' | Stock: '.$m['quantity'].' | Batch: '.$m['batch_number'].($ex?' | EXPIRED':''))?></option><?php endforeach; ?></select></div>
<div class="col-md-2"><label class="form-label">Quantity Out</label><input type="number" name="qty_out" class="form-control" min="1" value="1" required></div>
<div class="col-md-4"><label class="form-label">Patient / Recipient</label><input name="recipient" class="form-control" required></div>
<div class="col-12"><button class="btn btn-danger">Confirm Dispense</button></div></div></form></div>
<div class="cardx p-4"><h5 class="fw-bold">Dispensing History</h5><div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Date</th><th>Medicine</th><th>Batch</th><th>Qty Out</th><th>Recipient</th></tr></thead><tbody>
<?php foreach($dispenseLogs as $l): ?><tr><td><?=h($l['date'])?></td><td class="fw-bold"><?=h($l['inventory_name'])?></td><td><?=h($l['batch_number'])?></td><td class="text-danger fw-bold">-<?=$l['qty_out']?></td><td><?=h($l['recipient'])?></td></tr><?php endforeach; ?>
</tbody></table></div></div>
</section>

<section class="tab-pane fade" id="reports">
<div class="d-flex justify-content-between mb-3"><div><h4 class="fw-bold">Reports & Analytics</h4><small class="text-muted">Inventory and monthly dispensing reports.</small></div><a href="?export=inventory" class="btn btn-success">Inventory Excel</a></div>
<div class="row g-3 mb-4"><?php foreach([['Total Medicines',$totalProducts],['Total Units',number_format($totalStock)],['Expired',$expiredCount],['Expiring in 30 Days',$expiringCount]] as $r): ?><div class="col-md-3"><div class="cardx p-4"><small class="text-muted fw-bold"><?=$r[0]?></small><div class="num"><?=$r[1]?></div></div></div><?php endforeach; ?></div>
<div class="cardx p-4 mb-4"><h5 class="fw-bold">Monthly Medicines Dispensed</h5><form method="get" class="row g-2 mb-4">
<div class="col-md-3"><label class="form-label">Month</label><select name="report_month" class="form-select"><?php for($m=1;$m<=12;$m++): ?><option value="<?=$m?>" <?=$m===$selectedMonth?'selected':''?>><?=date('F',mktime(0,0,0,$m,1))?></option><?php endfor; ?></select></div>
<div class="col-md-2"><label class="form-label">Year</label><select name="report_year" class="form-select"><?php for($y=date('Y')-3;$y<=date('Y')+1;$y++): ?><option <?=$y===$selectedYear?'selected':''?>><?=$y?></option><?php endfor; ?></select></div>
<div class="col-md-3 d-flex align-items-end"><button class="btn btn-primary">Show Report</button></div>
<div class="col-md-4 d-flex align-items-end justify-content-end"><a class="btn btn-success" href="?export=dispensing&month=<?=$selectedMonth?>&year=<?=$selectedYear?>">Export This Month</a></div></form>
<div class="alert alert-danger"><strong><?=date('F Y',strtotime("$selectedYear-$selectedMonth-01"))?></strong> — Total dispensed: <strong><?=number_format($monthlyTotal)?> units</strong></div>
<h6 class="fw-bold">Medicines Released This Month</h6>
<?php if(!$monthlyMedicineTotals): ?><p class="text-muted">No medicines were dispensed during this month.</p><?php else: ?><div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Medicine</th><th>Total Released</th><th>Percentage</th></tr></thead><tbody>
<?php foreach($monthlyMedicineTotals as $name=>$q): $pct=$monthlyTotal?($q/$monthlyTotal*100):0; ?><tr><td class="fw-bold"><?=h($name)?></td><td class="text-danger fw-bold"><?=$q?> units</td><td><div class="progress"><div class="progress-bar bg-primary" style="width:<?=$pct?>%"><?=number_format($pct,1)?>%</div></div></td></tr><?php endforeach; ?>
</tbody></table></div><?php endif; ?>
<h6 class="fw-bold mt-4">Detailed Transactions</h6><div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Date</th><th>Medicine</th><th>Batch</th><th>Quantity</th><th>Recipient</th></tr></thead><tbody>
<?php if(!$monthlyDispense): ?><tr><td colspan="5" class="text-center text-muted">No transactions.</td></tr><?php else: foreach($monthlyDispense as $l): ?><tr><td><?=h($l['date'])?></td><td><?=h($l['inventory_name'])?></td><td><?=h($l['batch_number'])?></td><td class="text-danger fw-bold">-<?=$l['qty_out']?></td><td><?=h($l['recipient'])?></td></tr><?php endforeach; endif; ?>
</tbody></table></div></div>

<div class="cardx p-4"><h5 class="fw-bold">Expiration Report</h5><hr><h6 class="text-danger fw-bold">Already Expired</h6>
<?php if(!$expiredMedicines): ?><p class="text-muted">No expired medicines.</p><?php else: ?><div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Medicine</th><th>Batch</th><th>Expiration</th><th>Stock</th><th>Status</th></tr></thead><tbody><?php foreach($expiredMedicines as $m): ?><tr class="expired"><td><?=h(medicineFullName($m))?></td><td><?=h($m['batch_number'])?></td><td><?=h($m['expiration_date'])?></td><td>0</td><td><span class="badge bg-dark">EXPIRED</span></td></tr><?php endforeach; ?></tbody></table></div><?php endif; ?>
<hr><h6 class="text-warning fw-bold">Expiring Within 30 Days</h6>
<?php if(!$expiringMedicines): ?><p class="text-muted">No medicines expiring within 30 days.</p><?php else: ?><div class="table-responsive"><table class="table table-bordered"><thead><tr><th>Medicine</th><th>Batch</th><th>Expiration</th><th>Days</th><th>Stock</th></tr></thead><tbody><?php foreach($expiringMedicines as $m): $d=ceil((strtotime($m['expiration_date'])-$today)/86400); ?><tr class="expiring"><td><?=h(medicineFullName($m))?></td><td><?=h($m['batch_number'])?></td><td><?=h($m['expiration_date'])?></td><td><?=$d?></td><td><?=$m['quantity']?></td></tr><?php endforeach; ?></tbody></table></div><?php endif; ?>
</div>
</section>
</div>
</div>
</main>
</div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
Chart.defaults.font.family="Plus Jakarta Sans,system-ui,sans-serif";
new Chart(document.getElementById('trendChart'),{
 type:'line',
 data:{labels:<?=json_encode($trendLabels)?>,datasets:[{label:'Units Dispensed',data:<?=json_encode($trendData)?>,borderColor:'#7c3aed',backgroundColor:'rgba(124,58,237,.1)',fill:true,tension:.35}]},
 options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}}
});
new Chart(document.getElementById('categoryChart'),{
 type:'doughnut',
 data:{labels:<?=json_encode(array_keys($categoryCounts))?>,datasets:[{data:<?=json_encode(array_values($categoryCounts))?>,backgroundColor:['#7c3aed','#a78bfa','#c4b5fd','#ec4899','#f472b6','#5b21b6','#9061f9']}]},
 options:{cutout:'65%',plugins:{legend:{position:'bottom'}}}
});
new Chart(document.getElementById('statusChart'),{
 type:'doughnut',
 data:{labels:['Available','Low Stock','Expiring Soon','Expired'],datasets:[{data:[<?=$availableCount?>,<?=$lowStockCount?>,<?=$expiringCount?>,<?=$expiredCount?>],backgroundColor:['#16a34a','#f59e0b','#fb923c','#e11d48']}]},
 options:{cutout:'65%',plugins:{legend:{position:'bottom'}}}
});
</script>
</body>
</html>

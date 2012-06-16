#!c:/perl/bin/perl

use utf8;

use strict;
use warnings;
use Encode qw/encode decode/;
use DBI;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

my $DB_NAME = "resort_takeo";
my $DB_HOST = "localhost";
my $DB_USER = "postgres";
my $DB_PASS = "postgres";
my $dbh;

my %busho = ('客室' => 0, 'エトワール'=>1, '楠乃木'=>2, '桃花苑'=>3, 'ＮＹ－ＮＹ'=>4, 'ティファニー'=>5, '婚礼'=>6,
             '宴会'=>7, 'ＡＶスタジオ'=>8, 'コモンズ'=>9, 'ｲﾝﾌｫﾒｰｼｮﾝ'=>10, '慧洲園'=>11, 'その他'=>12 );

my $excel = Spreadsheet::ParseExcel->new;
my $fmt = Spreadsheet::ParseExcel::FmtJapan->new(Code=>'sjis');
my $book = $excel->parse("nipposample1.xls", $fmt);
my $sheet = $book->{"Worksheet"}[0];

# 部署名から部署コードを返す
sub getBushoCode {
    my $bushomei = shift;
    my @row_ary1 = $dbh->selectrow_array("select id from mt_busho where bushomei = '$bushomei'") or die $dbh->errstr;
    if (@row_ary1 != 0) {
        return $row_ary1[0];
    }
    
    my @row_ary2 = $dbh->selectrow_array("select id from mt_busho where bushomei2 = '$bushomei'") or die $dbh->errstr;
    if (@row_ary2 != 0) {
        return $row_ary2[0];
    }
    
    my @row_ary3 = $dbh->selectrow_array("select id from mt_busho where bushomei3 = '$bushomei'") or die $dbh->errstr;
    if (@row_ary3 != 0) {
        return $row_ary3[0];
    }
    
    return 0;
}

#　和暦(gg.m.d)を西暦に変換する関数
sub dateweststyle {
    my $gstyle = shift;
    my @p = split /\./, $gstyle;
    my $wstyle;
    if ($p[0] =~ m/([a-zA-Z]?)([0-9]*)/) {
        $wstyle = sprintf("%d", 1988 + $2) . '/' . sprintf("%d", $p[1]) . '/' . sprintf("%d", $p[2]);
    }
    return $wstyle;
}

# 客室売上をデータベースに保存
sub kyakushitsu {

#CREATE TABLE tr_kyakushitsu
#(
#  id serial NOT NULL,
#  uri_date date NOT NULL,
#  riyosu integer,
#  kadosu integer,
#  shitsuryo numeric(10,0) DEFAULT 0,
#  service numeric(10,0) DEFAULT 0,
#  telfee_taxab numeric(10,0) DEFAULT 0,
#  reizoko_taxab numeric(10,0) DEFAULT 0,
#  esthe_taxab numeric(10,0) DEFAULT 0,
#  other_taxab numeric(10,0) DEFAULT 0,
#  quo_taxfr numeric(10,0) DEFAULT 0,
#  cons_tax numeric(10,0) DEFAULT 0,
#  bath_tax numeric(10,0) DEFAULT 0,
#  shukugake numeric(10,0) DEFAULT 0,
#  genkin numeric(10,0) DEFAULT 0,
#  furikae numeric(10,0) DEFAULT 0,
#  credit numeric(10,0) DEFAULT 0,
#  coupon numeric(10,0) DEFAULT 0,
#  sotogake numeric(10,0) DEFAULT 0,
#  miseisan numeric(10,0) DEFAULT 0,
#  zenmiseisan numeric(10,0) DEFAULT 0,
#  miseisanzan numeric(10,0) DEFAULT 0,
#  CONSTRAINT tr_kyakushitsu_pkey PRIMARY KEY (id)
#)

#  H22.7.19～　客室附帯にQUOカード追加
#  H22.11.12～　客室附帯にWPC(World Prepaid Card)追加

    my @parm;
    my $datewareki = $sheet->{"Cells"}[19][1]->value(); # 日付
    # 和暦(gg.m.d)を西暦に変換する関数
    my $wstyle = dateweststyle($datewareki);
    push @parm, $wstyle;
    push @parm, $sheet->{"Cells"}[4][3]->unformatted();   # 利用人数
    push @parm, $sheet->{"Cells"}[21][4]->unformatted();  # 稼働室数
    push @parm, $sheet->{"Cells"}[4][4]->unformatted();   # 室料
    push @parm, $sheet->{"Cells"}[4][5]->unformatted();   # サービス料
    push @parm, $sheet->{"Cells"}[20][7]->unformatted();  # 客室附帯課税：電話代
    push @parm, $sheet->{"Cells"}[21][7]->unformatted();  # 客室附帯課税：冷蔵庫
    push @parm, $sheet->{"Cells"}[22][7]->unformatted();  # 客室附帯課税：エステ
    push @parm, $sheet->{"Cells"}[23][7]->unformatted();  # 客室附帯課税：その他
    push @parm, $sheet->{"Cells"}[25][7]->unformatted();  # 客室附帯非課税：QUOカード
    push @parm, $sheet->{"Cells"}[26][7]->unformatted();  # 客室附帯非課税：WPC
    push @parm, $sheet->{"Cells"}[4][20]->unformatted();  # 消費税
    push @parm, $sheet->{"Cells"}[4][21]->unformatted();  # 入湯税
    push @parm, $sheet->{"Cells"}[4][23]->unformatted();  # 宿掛
    push @parm, $sheet->{"Cells"}[4][25]->unformatted();  # 現金
    push @parm, $sheet->{"Cells"}[4][27]->unformatted();  # 振替
    push @parm, $sheet->{"Cells"}[4][28]->unformatted();  # クレジット
    push @parm, $sheet->{"Cells"}[4][29]->unformatted();  # クーポン
    push @parm, $sheet->{"Cells"}[4][30]->unformatted();  # 外掛
    push @parm, $sheet->{"Cells"}[4][31]->unformatted();  # 未精算額
    push @parm, $sheet->{"Cells"}[4][32]->unformatted();  # 前日未精算
    push @parm, $sheet->{"Cells"}[4][33]->unformatted();  # 未精算残高

    foreach my $pr (@parm) {
        if ($pr eq "") {
            $pr = 0;
        }
    }

    my $sql = "INSERT INTO tr_kyakushitsu(id, uri_date, riyosu, kadosu, shitsuryo, service, ".
        "telfee_taxab, reizoko_taxab, esthe_taxab, other_taxab, quo_taxfr, wpc_taxfr, cons_tax, ".
        "bath_tax, shukugake, genkin, furikae, credit, coupon, sotogake, miseisan, zenmiseisan, ".
        "miseisanzan) VALUES (default, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
    my $sth = $dbh->prepare($sql) or die $dbh->errstr;
    $sth->execute(@parm) or die $sth->errstr;
}

# 料飲売上（エトワール、Toukaen、ティファニー、宴会、婚礼、AV）をデータベースに保存
sub ryoin {

#CREATE TABLE tr_ryoin
#(
#  id serial NOT NULL,
#  uri_date date,
#  busho_code integer,
#  riyosu integer,
#  uriage_washoku numeric(10,0) DEFAULT 0,
#  uriage_chugoku numeric(10,0) DEFAULT 0,
#  uriage_yoshoku numeric(10,0) DEFAULT 0,
#  uriage_sonota numeric(10,0) DEFAULT 0,
#  uriage_inryo numeric(10,0) DEFAULT 0,
#  service numeric(10,0) DEFAULT 0,
#  uriage_futai numeric(10,0) DEFAULT 0,
#  uriage_waribiki numeric(10,0) DEFAULT 0,
#  cons_tax numeric(10,0) DEFAULT 0,
#  shukugake numeric(10,0) DEFAULT 0,
#  genkin numeric(10,0) DEFAULT 0,
#  furikae numeric(10,0) DEFAULT 0,
#  credit numeric(10,0) DEFAULT 0,
#  coupon numeric(10,0) DEFAULT 0,
#  sotogake numeric(10,0) DEFAULT 0,
#  CONSTRAINT tr_ryoin_pkey PRIMARY KEY (id)
#)

    my $bn = shift;

    my $riyosu = $sheet->{"Cells"}[4+$bn][3]->unformatted();   # 利用人数
    if ($riyosu eq "") {return;}
    
    my @parm;
    my $datewareki = $sheet->{"Cells"}[19][1]->value();
    # 和暦(gg.m.d)を西暦に変換する関数
    my $wstyle = dateweststyle($datewareki);

    push @parm, $wstyle;  # 日付
    push @parm, $bn;      # 部署コード
    push @parm, $riyosu;  # 利用人数
    push @parm, $sheet->{"Cells"}[4+$bn][8]->unformatted();   # 和食
    push @parm, $sheet->{"Cells"}[4+$bn][9]->unformatted();   # 中国食
    push @parm, $sheet->{"Cells"}[4+$bn][10]->unformatted();  # 洋食
    push @parm, $sheet->{"Cells"}[4+$bn][11]->unformatted();  # その他
    push @parm, $sheet->{"Cells"}[4+$bn][13]->unformatted();  # 飲料
    push @parm, $sheet->{"Cells"}[4+$bn][14]->unformatted();  # サービス料
    push @parm, $sheet->{"Cells"}[4+$bn][16]->unformatted();  # 料飲附帯
    push @parm, $sheet->{"Cells"}[4+$bn][19]->unformatted();  # 売上割引
    push @parm, $sheet->{"Cells"}[4+$bn][20]->unformatted();  # 消費税
    push @parm, $sheet->{"Cells"}[4+$bn][23]->unformatted();  # 宿掛
    push @parm, $sheet->{"Cells"}[4+$bn][25]->unformatted();  # 現金
    push @parm, $sheet->{"Cells"}[4+$bn][27]->unformatted();  # 振替
    push @parm, $sheet->{"Cells"}[4+$bn][28]->unformatted();  # クレジット
    push @parm, $sheet->{"Cells"}[4+$bn][29]->unformatted();  # クーポン
    push @parm, $sheet->{"Cells"}[4+$bn][30]->unformatted();  # 外掛

    foreach my $pr (@parm) {
        if ($pr eq "") {
            $pr = 0;
        }
    }

    my $sql = "INSERT INTO tr_ryoin(id, uri_date, busho_code, riyosu, uriage_washoku, uriage_chugoku, ".
        "uriage_yoshoku, uriage_sonota, uriage_inryo, service, uriage_futai, uriage_waribiki, cons_tax, ".
        "shukugake, genkin, furikae, credit, coupon, sotogake) VALUES ".
        "(default, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
    my $sth = $dbh->prepare($sql) or die $dbh->errstr;
    $sth->execute(@parm) or die $sth->errstr;
}

# その他売上（コモンズ、インフォメーション、その他）をデータベースに保存
sub sonota {

#CREATE TABLE tr_sonota
#(
#  id serial NOT NULL,
#  uri_date date,
#  busho_code integer,
#  riyosu integer,
#  uriage_sonota numeric(10,0) DEFAULT 0,
#  uriage_waribiki numeric(10,0) DEFAULT 0,
#  bath_tax numeric(10,0) DEFAULT 0,
#  shukugake numeric(10,0) DEFAULT 0,
#  genkin numeric(10,0) DEFAULT 0,
#  furikae numeric(10,0) DEFAULT 0,
#  credit numeric(10,0) DEFAULT 0,
#  coupon numeric(10,0) DEFAULT 0,
#  sotogake numeric(10,0) DEFAULT 0,
#  CONSTRAINT tr_sonota_pkey PRIMARY KEY (id)
#)

    my $bn = shift;

    my $riyosu = $sheet->{"Cells"}[4+$bn][3]->unformatted();   # 利用人数
    if ($riyosu eq "") {return;}
    
    my @parm;
    my $datewareki = $sheet->{"Cells"}[19][1]->value();
    # 和暦(gg.m.d)を西暦に変換する関数
    my $wstyle = dateweststyle($datewareki);

    push @parm, $wstyle;  # 日付
    push @parm, $bn;      # 部署コード
    push @parm, $riyosu;  # 利用人数
    push @parm, $sheet->{"Cells"}[4+$bn][17]->unformatted();  # その他
    push @parm, $sheet->{"Cells"}[4+$bn][19]->unformatted();  # 売上割引
    push @parm, $sheet->{"Cells"}[4+$bn][20]->unformatted();  # 消費税
    push @parm, $sheet->{"Cells"}[4+$bn][21]->unformatted();  # 入湯税
    push @parm, $sheet->{"Cells"}[4+$bn][23]->unformatted();  # 宿掛
    push @parm, $sheet->{"Cells"}[4+$bn][25]->unformatted();  # 現金
    push @parm, $sheet->{"Cells"}[4+$bn][27]->unformatted();  # 振替
    push @parm, $sheet->{"Cells"}[4+$bn][28]->unformatted();  # クレジット
    push @parm, $sheet->{"Cells"}[4+$bn][29]->unformatted();  # クーポン
    push @parm, $sheet->{"Cells"}[4+$bn][30]->unformatted();  # 外掛

    foreach my $pr (@parm) {
        if ($pr eq "") {
            $pr = 0;
        }
    }

    my $sql = "INSERT INTO tr_sonota(id, uri_date, busho_code, riyosu, uriage_sonota, uriage_waribiki, ".
        "cons_tax, ".
        "shukugake, genkin, furikae, credit, coupon, sotogake) VALUES ".
        "(default, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
    my $sth = $dbh->prepare($sql) or die $dbh->errstr;
    $sth->execute(@parm) or die $sth->errstr;

}


# ****** main ******
$dbh = DBI->connect("dbi:Pg:dbname=$DB_NAME;host=$DB_HOST", "$DB_USER", "$DB_PASS")
    or die "$!\nError: failed to connect to DB.\n";

kyakushitsu();
ryoin(1);  # エトワール
ryoin(3);  # 桃花苑
ryoin(5);  # ティファニー
ryoin(6);  # 婚礼
ryoin(7);  # 宴会
ryoin(8);  # AVスタジオ
sonota();

$dbh->disconnect;

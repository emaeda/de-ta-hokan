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

my $excel = Spreadsheet::ParseExcel->new;
my $fmt = Spreadsheet::ParseExcel::FmtJapan->new(Code=>'sjis');
my $book = $excel->parse("nipposample1.xls", $fmt);
my $sheet = $book->{"Worksheet"}[0];

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
    my $uri_date      = $sheet->{"Cells"}[19][1]->Value;  # 日付
    my $riyosu        = $sheet->{"Cells"}[4][3]->Value;   # 利用人数
    my $kadosu        = $sheet->{"Cells"}[21][4]->Value;  # 稼働室数
    my $shitsuryo     = $sheet->{"Cells"}[4][4]->Value;  # 室料
    my $service       = $sheet->{"Cells"}[4][5]->Value;  # サービス料
    my $telfee_taxab  = $sheet->{"Cells"}[20][7]->Value;  # 客室附帯課税：電話代
    my $reizoko_taxab = $sheet->{"Cells"}[21][7]->Value;  # 客室附帯課税：冷蔵庫
    my $esthe_taxab   = $sheet->{"Cells"}[22][7]->Value;  # 客室附帯課税：エステ
    my $other_taxab   = $sheet->{"Cells"}[23][7]->Value;  # 客室附帯課税：その他
    my $quo_taxfr     = $sheet->{"Cells"}[25][7]->Value;  # 客室附帯非課税：QUOカード
    my $wpc_taxfr     = $sheet->{"Cells"}[26][7]->Value;  # 客室附帯非課税：WPC
    my $cons_tax      = $sheet->{"Cells"}[4][20]->Value;  # 消費税
    my $bath_tax      = $sheet->{"Cells"}[4][21]->Value;  # 入湯税
    my $shukugake     = $sheet->{"Cells"}[4][23]->Value;  # 宿掛
    my $genkin        = $sheet->{"Cells"}[4][25]->Value;  # 現金
    my $furikae       = $sheet->{"Cells"}[4][27]->Value;  # 振替
    my $credit        = $sheet->{"Cells"}[4][28]->Value;  # クレジット
    my $coupon        = $sheet->{"Cells"}[4][29]->Value;  # クーポン
    my $sotogake      = $sheet->{"Cells"}[4][30]->Value;  # 外掛
    my $miseisan      = $sheet->{"Cells"}[4][31]->Value;  # 未精算額
    my $zenmiseisan   = $sheet->{"Cells"}[4][32]->Value;  # 前日未精算
    my $miseisanzan   = $sheet->{"Cells"}[4][33]->Value;  # 未精算残高
}

# 料飲売上（レストラン、ティファニー、宴会、婚礼）をデータベースに保存
sub ryoin {
    
}

# その他売上（コモンズ、インフォメーション、その他）をデータベースに保存
sub sonota {
    
}


# ****** main ******
$dbh = DBI->connect("dbi:Pg:dbname=$DB_NAME;host=$DB_HOST", "$DB_USER", "$DB_PASS")
    or die "$!\nError: failed to connect to DB.\n";

$dbh->disconnect;

kyakushitsu();
ryoin();
sonota();

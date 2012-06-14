#!c:/perl/bin/perl

use strict;
use utf8;
use Encode qw/encode decode/;

my $sqlterm1 = "insert into mt_shain(id, f_name, l_name, f_kana, l_kana, ".
               "find_name, birthday, nyusha, taishoku, postcode, jusho, ".
               "postcode2, jusho2, name2, acc_name, bank_code, bank_name, ".
               "branch_code, branch_name, acc_type, acc_number) values(";

# CSVファイルの読み込み
my $str1 = "読み込むCSVファイル名?:";
$str1 = encode("SJIS", $str1);
print $str1;
my $file = <STDIN>; # ファイル名読み込み 
chomp $file;

open my $fh , '<', $file
    or die "Cannot open '$file': $!";

while (my $line = <$fh>) {
    chomp $line;

    my @item = split(/,/, $line);

    my $sql = $sqlterm1;
    $sql.= @item[0] . ", ";         # 社員コード
    $sql.= "'" . @item[1] . "', ";  # 苗字
    $sql.= "'" . @item[2] . "', ";  # 名前
    $sql.= "'" . @item[3] . "', ";  # 苗字カナ
    $sql.= "'" . @item[4] . "', ";  # 名前カナ
    $sql.= "'" . @item[5] . "', ";  # 検索名
    $sql.= "'" . @item[6] . "', ";  # 生年月日
    $sql.= @item[8] eq 'NULL' ? @item[8].", " : "'".@item[8]."', ";  # 入社年月日
    $sql.= @item[9] eq 'NULL' ? @item[9].", " : "'".@item[9]."', ";  # 退職年月日
    $sql.= "'" . @item[10] . "', ";  # 郵便番号
    $sql.= "'" . @item[11] . "', ";  # 住所
    $sql.= @item[12] eq 'NULL' ? @item[12].", " : "'".@item[12]."', ";  # 郵便番号２
    $sql.= @item[13] eq 'NULL' ? @item[13].", " : "'".@item[13]."', ";  # 住所２
    $sql.= @item[15] eq 'NULL' ? @item[15].", " : "'".@item[15]."', ";  # 氏名２
    $sql.= "'" . @item[16] . "', ";  # 口座名義
    $sql.= "'" . sprintf("%04d", @item[17]) . "', ";  # 銀行コード
    $sql.= "'" . @item[18] . "', ";  # 銀行名
    $sql.= "'" . sprintf("%03d", @item[19]) . "', ";  # 支店コード
    $sql.= "'" . @item[20] . "', ";  # 支店名
    $sql.= @item[21] . ", ";  # 区分
    $sql.= "'" . sprintf("%07d", @item[22]) . "'";    # 口座番号
    $sql.= ");";  # SQL文の終わり
    print $sql."\n";
}

use strict;
use warnings;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;


package Excellib;

# in: Spalten-Zahl
# out: Excel->Rangel-Objekt?
# oder string
# TODO A1-Format via Excel-Funktion
sub rangetocell_format {
    my $col = shift;
    my %rangetocell_lib = (
       1=>"A",
       2=>"B",
       3=>"C",
       4=>"D",
       5=>"E",
       6=>"F",
       7=>"G",
       8=>"H",
       9=>"I",
       10=>"J",
       11=>"K",
       12=>"L",
       13=>"M",
       14=>"N",
       15=>"O",
       16=>"P",
       17=>"Q",
       18=>"R",
       19=>"S",
       20=>"T",
       21=>"U",
       22=>"V",
       23=>"W",
       24=>"X",
       25=>"Y",
       26=>"Z",
       #27=>"AA",
    );
    my $col_A1 = '';
    while ($col >= 27) {
        my $col_temp = $col-26;
        $col = $col-26;
        $col_A1 .= $rangetocell_lib{$col_temp};
        }
    $col_A1 .= $rangetocell_lib{$col};
    return $col_A1;
}

1;
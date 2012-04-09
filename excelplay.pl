use strict;
use warnings;
use excel_com2;
use feature qw/say/;

# get active OLE automation objects: program or class id


#my $excelfile = "f:\\poly\\HU-tp4\\2012-03-15 Vereinsliste aktive.xlsx";
my $excelobj = Excelobject->new();
#$excelobj->init($excelfile,4);
$excelobj->init();
# $excelobj->Zeilen_in_1Spalte(2,6);	# find last cell funktioniert nicht
$excelobj->Zeilen_in_1Spalte(2,1,30,1);

#########
## alt ##
#my $colstart = 1;
#my $srow = 23;
#my $scol = 1;
#my @vals;
#while (my $vals_ref = [ $excelobj->readcol($colstart, 2) ] ) {
#	say $colstart.",";
#	my @vals = @{$vals_ref};
#	last if (scalar @vals) == 0;
#	foreach my $val (@vals) {
#		$excelobj->pos($srow, $scol);
#		$excelobj->writeval($val);
#		$srow++;
#	}
#	$colstart++;
#}




print "__End\n";
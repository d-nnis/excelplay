use strict;
use warnings;
use excel_com;

#my $excelfile = "f:\\tee.xlsx";
#my $excelobj=Excelobject->new();
#$excelobj->init_newExcelfile($excelfile, 1);
##$excelobj->init_Excelfile();
#
#my @vals = $excelobj->readcol(2);
#print @vals;


# get active OLE automation objects: program or class id


my $excelfile = "f:\\poly\\HU-tp4\\2012-03-15 Vereinsliste aktive.xlsx";
my $excelobj = Excelobject->new() ;
$excelobj->init_newExcelfile($excelfile,4);
my $colstart = 2;
my $srow = 23;
my $scol = 1;
my @vals;
while (my $vals_ref = [ $excelobj->readcol($colstart, 2) ] ) {
	$|=1;
	print $colstart.",";
	$|=1;
	my @vals = @{$vals_ref};
	last if (scalar @vals) == 0;
	foreach my $val (@vals) {
		$excelobj->pos($srow, $scol);
		$excelobj->writeval($val);
		$srow++;
	}
	$colstart++;
}




print "__End\n";
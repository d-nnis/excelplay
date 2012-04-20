use strict;
use warnings;
use excel_com2;
use feature qw/say/;

# get active OLE automation objects: program or class id
print "regexcell-HEAD\n";

#tee
#my $excelfile = "f:\\poly\\HU-tp4\\2012-03-15 Vereinsliste aktive.xlsx";
my $excelobj = Excelobject->new();
#$excelobj->init($excelfile,4);
$excelobj->init();
#$excelobj->Zeilen_in_1Spalte(2,2,10,2);

# transpose_level 0: Formelbezug
# 1: Wert kopieren

#$excelobj->active_cell('aim');
$excelobj->transpose_level(1);
#print "transpose_level:",$excelobj->transpose_level,"\n";
#$excelobj->Zeilen_in_1Spalte(2,2,10,1);
# TODO
$excelobj->regex('addcell');
#$excelobj->{regexp} = '(\d)(\d)';
$excelobj->{regexp} = 'activecell';

my $range = $excelobj->regex_col(4,11);


print "ende\n";
__END__

## merge cells
#$excelobj->join_row(2,14);
$excelobj->set_join_sep('-');
$excelobj->join_row_block(2,14);

##

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
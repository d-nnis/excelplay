
use feature "switch";
use File::Copy;
use strict;
use warnings;
# popri
#use lib('C:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\Eigene Dateien\\workspace\\V6\\');
# home
#use lib('C:\\Documents and Settings\\dennis\\My Documents\\workspace\\PMs\\');
#use lib('');
use Essent;
use Excel_com;
use db_com;
use sps_handle;
use v6_settings;
use DBI;
#use diagnostics;
#enable diagnostics;


########################
# db_structure
#
# db-struktur an legen thnum -> ver, ohne Daten von FORMS !
#
#
# für jedes sps file:
# lies aus
# ergänze Einträge in xls-File um das SPSS-Format
# todo xl-Ergänzung
#   (1)
#   sps-Files pro TH-set zusammen fassen -> xl-sheet!
#   sps-Files auf Existenz prüfen (-e, wenn nicht, aufschreien)
#   und gegen xl-sheet laufen lassen (complete_sheet)
#
#   (2)
#   --> alle Einträge ergänzt?
#    -> dann xls-File auslesen und Struktur in db anlegen, inkl spss-format


# next step: db_structure.pl

#####################
# INPUT
# a) sps-files eines sets (TH), bzw. Excel-Def-Sheet als Array (lookup)
#   und Files zum einlesen


my @Chwa_V1_defs = qw(
Chwa-S03.SPS
Chwa-S04.SPS
Chwa-S05.SPS
Chwa-S06.SPS
Chwa-S07.SPS
Chwa-S08.SPS
Chwa-S09.SPS
Chwa-S10.SPS
Chwa-S11.SPS
Chwa-S12.SPS
Chwa-S13.SPS
Chwa-S14.SPS
Chwa-S15.SPS
Chwa-S16.SPS
);

my @MSR1 = qw(
MSR1_R.SPS
MSR1_V.SPS
);

my @MSR2 = qw(
MSR2_R.SPS
MSR2_V.SPS
);

my @MSR3 = qw(
MSR3_R.SPS
MSR3_V.SPS
);

my @defset = @MSR3;
#my $readdir = Popri::Chwa::transfer();
# popri
#my $readdir = 'C:\\forms\\projekt\\Chwa-V2\\transfer\\';
# home
# sps-files
my $readdir = 'C:\\forms\\projekt\\MSR2\\transfer\\';
my $exceldir = 'C:\\forms\\projekt\\MSR2\\transfer\\';
my $excelfile = join('', $exceldir, 'MSR.xls');
my $excelsheet = 3;

check_existence(@defset);

print "Ergaenze Excel-File: ", $excelfile, " SheetNo: ", $excelsheet, ".\n";

# in welche excel-datei schauen?
# Chwa-V1
my $complete_excel = Complete_Excel->new();
# excelfile, excelsheet
$complete_excel->init_Excelfile($excelfile, $excelsheet);

# xl-sheet mit SPS-Format und Importfi ergänzen
#
foreach my $spsfile (@defset) {
	my $spsfile = join('', $readdir, $spsfile);
	print "Abgleich mit ", $spsfile, "...\n";
	my $spsobj =  spsfile->new('no_order');
	my $vars_ref = $spsobj -> varHash(File::readfile($spsfile));
	my %vars = %{$vars_ref}; 
	#print "keys of spsfile:", keys %vars, "\n";
	# übergib vars aus sps-file in spss-format
	$complete_excel->var_format_spss({%vars});
	$complete_excel->complete_sheet();
}


print "\nend\n";
print "...allet in xl-Sheet drin?\n";



sub check_existence {
	my @defs = @_;
	my $exit_flag = 0;
	foreach my $def (@defs) {
		if (-e join('', $readdir, $def) ) {
			print "SPS-File: ", join('', $readdir, $def), " existiert\n";
		} else {
			print "!!!!!!\n";
			print "SPS-File: ", join('', $readdir, $def), " existiert nicht!\n";
			print "!!!!!!\n";
			$exit_flag = 1;
		}
	}
	die 'Ueberpruefe alle sps-files!' if ($exit_flag);
	print "\nAlle SPS-Files existieren.\n";
}

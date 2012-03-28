
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
use sps_handle;

use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;

#use diagnostics #-verbose;
#enable diagnostics;
# Excel-part



##############################
# present_data.pl
#
# Generiere Excel-Sheet
# 
# input:
# Datenformat (sps)
# Datensatz
# Bilddateien
#
# 1st row: Variablennamen (aus sps)
#
# each row einfügen:
#   Datensatzzeile
#   relativer softlink zur Bilddatei
#   thumbnail (Zeile vergrößern)
# 
# deps:
#  spss-files
#  bildd_dir
#  bildd/ Importfi


# input-files
#my $dir = 'C:\\Documents and Settings\\dennis\\My Documents\\workspace\\V6\\present_data\\';
#my $dir = 'C:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\workspace\\V6\\export\\';
my $dir = 'c:\\forms\\projekt\\MSR2\\transfer\\';
my $spsfile = join('', $dir, 'MSR2.sps');
my $datfile = join('', $dir, 'MSR2.dat');

# output-files
my $excelfile = join('', $dir, 'MSR_2.xls');

# bildd-dir
my $bildd_rel = 'bildd\\';
my $bildd_dir = join('', 'C:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\workspace\\V6\\export\\bildd\\');
my $bildd_dir_nest = 'c:\\forms\\projekt\\MSR2\\transfer\\bildd\\';
my $schnipp_rel = 'schnipp\\';

my $spsobj = spsfile->new('order');

# Hash mit Referenzen zu Hashs
my $ordered_vars_ref = $spsobj->varHash(File::readfile($spsfile)); 

# SPSS-Daten einlesen
# '  112345               26C05000.TIF'
my @data = File::readfile($datfile);

# create split_image.bat
# nimm Importfi 'C:\forms\projekt\MSR2\import\2\identif.\FR713_V2.p00000019.tif'
# verkürze auf 'FR713_V2.p00000019.tif'
# i_view32.exe FR713_V2.p00000019.tif /crop=(1,1,2474,1800) /convert=FR713_V2.p00000019_kennw.png
# setze thumbnail (nested) mit _kennw erweitert ein: 'FR713_V2.p00000019_kennw.png'
#


print "xl-file.", $excelfile, "\n";
if (-e $excelfile) {print "exist\n";} else { print "NOT exist\n";}
#Process::confirm;
my $excelobj = Excelobject->new();
$excelobj->init_newExcelfile($excelfile, 1);
# bildd_dirs(link-dir, nest-dir)
$excelobj->bildd_dirs('', $bildd_dir_nest);

# übertrage Referenzen-Hash von Hashes auf
# (a) Array mit keys und
# (b) echtem Hash 
my @var_head;
my %kla_vars;
# key (spsvar) -> val (spsformat)
# key{'FrB01'} = '9-13';
for (my $j = 0; $j < scalar keys %{$ordered_vars_ref}; $j++) {
	(my $key, my $val) = %{$ordered_vars_ref->{$j}};
	push (@var_head, $key);
	$kla_vars{$key} = $val;
}

#print "varhead:", @var_head, "\n";
#Process::confirm;
# Schreibe Variablen-Überschriften in excel-sheet
$excelobj->add_head(@var_head);
# übergib einteilung %kla_vars in substr-format
$excelobj->var_format_tosubstr({%kla_vars});
$excelobj->adjust_cells();

# gehe jeden case durch (Datenzeile)
for (my $i = 0; $i < scalar @data; $i++) {
	print "insert case ", $i+1, "\n";
	$excelobj->insert_case($i+2, $data[$i]);
}


print "ende\n";

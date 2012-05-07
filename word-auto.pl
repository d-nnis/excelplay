use strict;
use warnings;
use Essent;
use Win32::OLE;
#use Win32::OLE::Const ('Microsoft Excel 8.0 Object Library');
#use Win32::OLE::Const('Microsoft Word');
use Win32::OLE::Const 'Microsoft Word';

my $path = "z:\\Lernstand_VERA 6 - 2012\\Auswertungsanleitungen\\";
my @files = File::get_by_ext($path, "doc[x]?");

### open Word application and add an empty document
### (will die if Word not installed on your machine)
my $word = Win32::OLE->new('Word.Application', 'Quit') or die;
$word->{Visible} = 1;

#my $doc = $word->Documents->Add();
my $doc = $word->Documents->Open($path.$files[0]);
## Layout
#ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
#$doc->PageSetup->{'TopMargin'} = 'CentimetersToPoints(1)';
$doc->PageSetup->{'TopMargin'} = 10;
$doc->PageSetup->{'BottomMargin'} = 10;
$doc->PageSetup->{'LeftMargin'} = 10;
$doc->PageSetup->{'RightMargin'} = 10;
## /Layout

#my $x = $word->Documents->ReadingLayoutSizeX;
#$self->{WORKSHEET}->Shapes->AddPicture (
#	$schnipp_abs,		# Filename As String
#	1,					# LinkToFile As MsoTriState
#	1,					# SaveWithDocument As MsoTriState
#	490,				# Left As Single
#	$top_align,	# Top As Single
#	350,				# Width As Single
#	40					# Height As Single
#);





my $range = $doc->{Content};

### insert some text into the document
$range->{Text} = 'Hello World from Monastery.';
$range->InsertParagraphAfter();
$range->InsertAfter('Bye for now.');

### read text from the document and print to the console
my $paras = $doc->Paragraphs;
foreach my $para (in $paras) {
	print ">> " . $para->Range->{Text};
}

### close the document and the application
$doc->SaveAs(FileName => "c:\\temp\\temp.txt", FileFormat => wdFormatDocument);
$doc->Close();
$word->Quit();


#my $word = init();


sub init {
	#my $self = shift;
	my $word;
	# TODO: test, wenn keine Excel-Instanz laeuft etc.
	# TODO: bestimmtes Excel-File öffnen
	# TODO: alle Excel-Threads erfassen und aufzählen/ wählen, CREATOR?
	# You can also directly attach your program to an already running OLE server:
	print "Count OLE-Objects:", Win32::OLE->EnumAllObjects(),"\n";
	eval {$word = Win32::OLE->GetActiveObject('Word.Application')};
	die "Word not installed" if $@;
	# create new word-server, returns a server object
	unless (defined $word) {
		$word = Win32::OLE->new('Word.Application', 'Quit')
				or die "Cannot start Word";
	}
	# Word-Server
	#$self->{WORD} = $word;
	
	## suche window aus?
	#my $window_count = $excel->Windows->Count;
	#if ($window_count == 0) {
	#	#$excel->Windows->Add;
	#	$excel->Windows(1)->Open;
	#	#$excel->{Visible} = 1;
	#	
	#} else {
	#	print "$window_count Windows: ";
	#	foreach (1..$window_count) {
	#		print "$_:".$excel->Windows($_)->Parent->Name.",";
	#		# oder: $ex->ActiveWindow->Caption;
	#	}
	#	print "\n";
	#	my $parent = $excel->Windows(1)->Parent->Name;
	#	my $active2 = $excel->ActiveWorkbook;
	#	print "";
	#}
	## suche workbook aus
	
	#my $workb_count = $excel->Workbooks->Count;
	#if ($workb_count == 0) {
	#	$excel->Workbooks->Add;
	#} elsif ($workb_count == 1) {
	#	$excel->Workbooks(1)->Activate;
	#	print "select workbook: '", $excel->Workbooks(1)->Name, "'\n";
	#	$self->{WORKBOOK} = $excel->Workbooks(1);
	#} else {
	#	print "$workb_count Workbooks: ";
	#	foreach (1..$workb_count) {
	#		print "$_:".$excel->Workbooks($_)->Name.",";
	#	}
	#	print "\n";
	#	my $workb_select = Process::confirm_numcount($workb_count);
	#	$excel->Workbooks($workb_select)->Activate;
	#	print "select workbook: '", $excel->Workbooks($workb_select)->Name, "'\n";
	#	$self->{WORKBOOK} = $excel->Workbooks($workb_select);
	#}
	#
	###
	### suche worksheet aus
	#my $works_count = $excel->Worksheets->Count;
	#print "$works_count Sheets: ";
	#if ($works_count == 1) {
	#	$self->{WORKSHEET} = $excel->Worksheets(1);
	#	print "select worksheet: '", $excel->Worksheets(1)->Name, "'\n";
	#} else {
	#	foreach (1 .. $works_count) {
	#		print "$_:".$excel->Worksheets($_)->Name.",";
	#	}
	#	print "\n";
	#	my $works_select = Process::confirm_numcount($works_count);
	#	## TODO: ausbauen um mit mehreren Sheets zu arbeiten
	#	$self->{WORKSHEET} = $excel->Worksheets($works_select);
	#	print "select worksheet: '", $excel->Worksheets($works_select)->Name, "'\n";
	#}
	return $word;
}

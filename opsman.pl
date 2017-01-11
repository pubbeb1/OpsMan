###TODO: Add country prio to follow up


use Win32::OLE;
use Win32::OLE qw(in with);
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft Excel';
use Time::Local;
use Time::localtime;
use Time::gmtime;


sub seekFolder {
  $obj = shift;
  $target = shift;

  for ($i = 1; $i <= $obj->Folders->Count; $i++) {
    if ( $obj->Folders->Item($i)->Name =~ /$target/ ) {
      return $obj->Folders->Item($i);
    }
  }
}

#$testing = "_testing";
$script_location = "\\\\ent\\mit-mst02\\EOC\\PMS\\Department\\Metrics_Spotfire\\";
$script_filename = "Daily WIP Report PMS-IW R4 v4.8.xlsx";
#$script_filename = "Daily WIP Report PMS-IW R4 v4.8_testing.xlsx";

$opsnames = "Brigitte Froels,Erik Coenen,Jeroen Saccol,Samantha Schmeets,Mirijan Schuurmans,Boris Pubben,Maurice Jacobs,Mara Honings,Danielle Smeets,Sascha Winkelmolen,Marjo Lintjens,Isabeau Joosten,Nathalie Pauly,UNASSIGNED UNASSIGNED";
$priocountries = "Russian Federation,Czech Republic";

$excel_max_row = 1048576;

$prio_color = 40;
$update_color = 3;
$alert_color = 3;

$header1_color = 35;
$header2_color = 36;

$do_completions = "TRUE";
#$do_completions = "FALSE";

#$tm = localtime(time);
$tm = gmtime(time);
($current_sec, $current_min, $current_hour, $current_day, $current_month, $current_year) = (sprintf("%02d",$tm->sec), sprintf("%02d",$tm->min), sprintf("%02d",$tm->hour), sprintf("%02d",$tm->mday), sprintf("%02d",($tm->mon + 1)), sprintf("%02d",($tm->year + 1900)));
my @wday = qw/Mon Tue Wed Thu Fri Sat Sun/;
$weekday = $wday[$tm->wday -1];

eval {$outlook = Win32::OLE->GetActiveObject('Outlook.Application')};
   die "Outlook not installed" if $@;
   unless (defined $outlook) {
      $outlook = Win32::OLE->new('Outlook.Application', sub {$_[0]->Quit;})
      or die "Oops, cannot start Outlook";
   }
   
# get already active Excel application or open new
my $Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
my $Book = $Excel->Workbooks->Open($script_location.$script_filename);

#my $Statistics = $Excel->Workbooks->Open($script_location."opsman\\opsman_statistics.xlsx");
my $WipBook = $Excel->Workbooks->Add;
if (not $Book || not $WipBook)
{
  die "Excel file not open\n";
}

#setup worksheets in order
$WipBook->worksheets(3)->Delete;
$WipBook->worksheets(2)->Delete;
$coversheet = $WipBook->worksheets(1);
$coversheet->{Name} = "Cover Sheet";
$sr_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("Cover Sheet")});
$sr_sheet->{Name} = "Staging Records";
$intakes_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("Staging Records")});
$intakes_sheet->{Name} = "Intakes";
$pe_updates_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("Intakes")});
$pe_updates_sheet->{Name} = "PE Updates";
$followup_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("PE Updates")});
$followup_sheet->{Name} = "Follow Up";
$sr_comms_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("Follow Up")});
$sr_comms_sheet->{Name} = "SR Communications";
$completions_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("SR Communications")});
$completions_sheet->{Name} = "Completions";
$qc_completions_sheet = $WipBook->Worksheets->Add({After=>$WipBook->Worksheets("Completions")});
$qc_completions_sheet->{Name} = "Quality Check - Completions";
$coversheet->Activate();


$followup_gch_list = "";


foreach $sheet (in $Book->worksheets)
{

  $sheetname = $sheet->{Name};

  if ($sheetname eq "Staging Records")
  {
  
    $staging_record_count = 0;
    $staging_record_urgent_count = 0;
    $staging_record_sr_count = 0;
    $staging_record_sr_urgent_count = 0;
    
    $line = 2;
  
  
    print "Collecting + Sorting Staging Records\n";
    
    @staging_records_prio = ();
    @staging_records_no_prio = ();
    @staging_records_dnp = ();
    
    while ($gch_number = $sheet->Range("A$line")->{Value})
    {
      $bu = $sheet->Range("H$line")->{Value};
      $patient_status = $sheet->Range("D$line")->{Value};
      $description = $sheet->Range("C$line")->{Value};
      $country = $sheet->Range("I$line")->{Value};
      
      if ($bu eq "Spinal" || $bu eq "NV" || $bu =~ /^CV-/ || $bu eq "NEUROMOD" || $bu eq "Cryocath" || $bu eq "AFI" || $bu eq "AFS" || $bu eq "Xomed" || $bu eq "NeuroSurgery" || $bu eq "PSS" || $bu eq "MAE") { $prio = -1; }

      elsif ($description =~ /DO NOT PROMOTE/i)
      {
        $prio = 3;
        $arrayref = \@staging_records_dnp;
      }
      elsif ($patient_status ne "" && $patient_status ne "Alive" || index($priocountries, $country) != -1 || $bu eq "")
      {
        $prio = 1;
        $arrayref = \@staging_records_prio;
      }
      else
      {
        $prio = 2;
        $arrayref = \@staging_records_no_prio;
      }

      if ($prio > 0)
      {
        my %prev_sr = %{@$arrayref[-1]};
        if ($prev_sr{gch_number} eq $gch_number)
        {	
          $prev_sr{bu} = $prev_sr{bu}."; ".$bu;
          pop @$arrayref;
          push @$arrayref,\%prev_sr;
        }
        else 
        {
          my %staging_record = (
            gch_number => $gch_number,
            bu => $bu,
            patient_status => $patient_status,
            description => $description,
            country => $country,
            notified_date => (join '-', reverse split '-', $sheet->Range("F$line")->{Value}),
            created_date => (join '-', reverse split '-', $sheet->Range("O$line")->{Value}),
            source => $sheet->Range("J$line")->{Value},
            source_id => $sheet->Range("K$line")->{Value},
            prio => $prio,
          );
          push @$arrayref,\%staging_record;			
        }
      }
      $line++;
    }
    
    
    ### Create SR destination worksheet and setup staging records header
    
    $sr_sheet_line = 1;
    
    $sr_sheet->Range("A$sr_sheet_line")->{value} = $sheetname;
    $sr_sheet->Range("A$sr_sheet_line")->Interior->{ColorIndex} = 35;
    $sr_sheet_line++;
    $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->{value} = ["PE#", "Aging", "Notified Date", "Created Date", "BU", "Country", "Patient Status", "Description", "Source", "Source ID"];
    $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->Interior->{ColorIndex} = 36;
    $sr_sheet_line++;
  
    ### OUTPUT Staging Records
    print "Adding Prio and normal WIP staging records to sheet\n";
    foreach (@staging_records_prio,@staging_records_no_prio)
    {
        my %sr_hash = %{$_};
        $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->{value} = [$sr_hash{gch_number}, "=today()-c$sr_sheet_line", $sr_hash{notified_date}, $sr_hash{created_date}, $sr_hash{bu}, $sr_hash{country}, $sr_hash{patient_status}, $sr_hash{description}, $sr_hash{source}, "\'$sr_hash{source_id}"];
        $sr_sheet->Range("C$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
        $sr_sheet->Range("D$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
        $sr_sheet->Range("B$sr_sheet_line")->{NumberFormat} = "####";
        #convert formula to value as workaround for timezone differences in google docs
        $testdate = $sr_sheet->Range("B$sr_sheet_line")->{Value};
        $sr_sheet->Range("B$sr_sheet_line")->{Value} = $testdate;
        $urgent = $sr_hash{prio};
        
        if ($weekday eq "Mon" || $weekday eq "Tue" ) { $compareDays = 4 } else { $compareDays = 2 }
        
        
        if ($sr_sheet->Range("B$sr_sheet_line:B$sr_sheet_line")->{value} > $compareDays)
        {
          #Disabled aging coloring because of new way of working
          #$sr_sheet->Range("B$sr_sheet_line")->Interior->{ColorIndex} = $prio_color;
          $urgent = 1;
        }
        
        if ($sr_hash{patient_status} ne "" && $sr_hash{patient_status} ne "Alive") 
        { 
          $sr_sheet->Range("A$sr_sheet_line,G$sr_sheet_line")->Interior->{ColorIndex} = $prio_color;
          $urgent = 1;          
        }
        elsif (index($priocountries, $sr_hash{country}) != -1)
        { 
          $sr_sheet->Range("A$sr_sheet_line,F$sr_sheet_line")->Interior->{ColorIndex} = $prio_color; 
          $urgent = 1;
        }
        elsif ($sr_hash{bu} eq "")
        {
          $sr_sheet->Range("A$sr_sheet_line,E$sr_sheet_line")->Interior->{ColorIndex} = $prio_color; 
          $urgent = 1;
        }
        
         if ($sr_hash{bu} eq "MAE" && $sr_hash{source} eq "SAP ECC Service and Repair")
        { 
          $sr_sheet->Range("A$sr_sheet_line,E$sr_sheet_line")->Interior->{ColorIndex} = 3; 
          $sr_sheet->Range("A$sr_sheet_line,I$sr_sheet_line")->Interior->{ColorIndex} = 3; 
        }
            
        $sr_hash{source} eq "SAP ECC Service and Repair" ? $staging_record_sr_count ++ : $staging_record_count++;
        if ($urgent == 1) { $sr_hash{source} eq "SAP ECC Service and Repair" ? $staging_record_sr_urgent_count ++ : $staging_record_urgent_count++; }		 
        $sr_sheet_line++;
    }	
  
    #PIRs to Process
    $sr_sheet_line++;
    $sr_sheet->Range("A$sr_sheet_line")->{value} = "PIRs to process";
    $sr_sheet->Range("A$sr_sheet_line")->Interior->{ColorIndex} = 35;
    $sr_sheet_line++;
    $sr_sheet->Range("A$sr_sheet_line:D$sr_sheet_line")->{value} = ["Subject", "Aging", "Received data", "Flag"];
    $sr_sheet->Range("A$sr_sheet_line:D$sr_sheet_line")->Interior->{ColorIndex} = 36;
    $pirs_to_process = 0;
    $sr_sheet_line++;
    
  
    $mailbox = seekFolder($outlook->Session, 'RS PP EMEA');
    $inbox = seekFolder($mailbox, 'Inbox');
    $PIRs_to_process = seekFolder($inbox, 'SR Creation');
    $folder = $PIRs_to_process->Items;
    $pirs_to_process_count = 0;
    $pirs_to_process_urgent_count = 0;
    print "Collecting PIRs to process\n";
    for ($i = $PIRs_to_process->Items->Count ; $i >=1 ; $i--)
    {
      $item = $folder->Item($i);
      $subject = substr($item->Subject,0,20);
      $sent = $item->CreationTime;
      $sent = join '-', reverse split '-', substr($sent, 0, index($sent, " "));
      $bu = $item->Categories;
      if (!$bu) {$bu = "Unassigned"}
      $sr_sheet->Range("A$sr_sheet_line:D$sr_sheet_line")->{value} = [$subject, "=today()-C$sr_sheet_line", $sent, $bu];
      $sr_sheet->Range("C$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
      $sr_sheet->Range("B$sr_sheet_line")->{NumberFormat} = "####";
      $pirs_to_process_count++;
      if ($weekday eq "Mon" || $weekday eq "Tue" ) { $compareDays = 4 } else { $compareDays = 2 }
      if ($sr_sheet->Range("B$sr_sheet_line:B$sr_sheet_line")->{value} > $compareDays)
      {
        $pirs_to_process_urgent_count++;
        #disabled aging coloring due to new way of working
        #$sr_sheet->Range("B$sr_sheet_line")->Interior->{ColorIndex} = $prio_color;
      }
      $sr_sheet_line++;
    }
    $sr_sheet_line++;
    
    ## Output DO NOT PROMOTE staging records
    
    print "Adding DO NOT PROMOTE staging records to sheet\n";
    
    $sr_sheet->Range("A$sr_sheet_line")->{value} = "Do not promote";
    $sr_sheet->Range("A$sr_sheet_line")->Interior->{ColorIndex} = 35;
    $sr_sheet_line++;
    $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->{value} = ["SR#", "Aging", "Notified Date", "Created Date", "BU", "Country", "Patient Status", "Description", "Source", "Source ID"];
    $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->Interior->{ColorIndex} = 36;
    $sr_sheet_line++;
    $line = 2;
      
    foreach (@staging_records_dnp)
    {
      my %sr_hash = %{$_};
      $sr_sheet->Range("A$sr_sheet_line:J$sr_sheet_line")->{value} = [$sr_hash{gch_number}, "=today()-c$sr_sheet_line", $sr_hash{notified_date}, $sr_hash{created_date}, $sr_hash{bu}, $sr_hash{country}, $sr_hash{patient_status}, $sr_hash{description}, $sr_hash{source}, "\'$sr_hash{source_id}"];
      $sr_sheet->Range("C$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
      $sr_sheet->Range("D$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
      $sr_sheet->Range("B$sr_sheet_line")->{NumberFormat} = "#0";
      #convert formula to value as workaround for timezone differences in google docs
      $testdate = $sr_sheet->Range("B$sr_sheet_line")->{Value};
      $sr_sheet->Range("B$sr_sheet_line")->{Value} = $testdate;
      
      $urgent = $sr_hash{prio};
      
      if ($weekday eq "Mon" || $weekday eq "Tue" ) { $compareDays = 4 } else { $compareDays = 2 }
      
      if ($sr_sheet->Range("B$sr_sheet_line:B$sr_sheet_line")->{value} > $compareDays)
      {
        #disabled aging coloring because of new way of working
        #$sr_sheet->Range("B$sr_sheet_line")->Interior->{ColorIndex} = $prio_color;
        $urgent = 1;
      }
          
      if ($sr_hash{patient_status} ne "" && $sr_hash{patient_status} ne "Alive") 
      { 
        $sr_sheet->Range("A$sr_sheet_line:B$sr_sheet_line,G$sr_sheet_line")->Interior->{ColorIndex} = $prio_color; 
      }
      elsif (index($priocountries, $sr_hash{country}) != -1)
      { 
        $sr_sheet->Range("F$sr_sheet_line")->Interior->{ColorIndex} = $prio_color; 
      }
      
      #$staging_record_count++;
      #if ($urgent == 1) { $staging_record_urgent_count++; }		 
      $sr_sheet_line++;
    }	
    
    ### Layout update
    $sr_sheet->Columns("A:XX")->{AutoFit}= "True";    
      
  }

   
  if ($sheetname eq "Follow Up")
  {
    print "Collecting Follow up\n";
    $line = 2;
    
    ##pre-sort on due date
    $sheet->Range("A$line:V$excel_max_row")->Sort({Key1 => $sheet->Range("I$line:I$excel_max_row"), Order1 => xlAscending});
    
    do
    {
      $com_id = $sheet->Range("E$line")->{Value};
      if ($com_id)
      {
        $gch_number = $sheet->Range("A$line")->{Value};
        
        $geo_person = $sheet->Range("C$line")->{Value};
        $description = $sheet->Range("H$line")->{Value};
        $comm_person_resp = $sheet->Range("K$line")->{Value};
        $country = $sheet->Range("D$line")->{Value};
        $due_date = $sheet->Range("I$line")->{Value};
        
        $bu = $sheet->Range("B$line")->{Value};
        $bu_filter = "FALSE";
                
        
        #20160725 (BP) Filter disabled as everyone should be able to handle intakes during follow up
        #if (($bu eq "NEUROMOD" || $bu eq "Cryocath" || $bu eq "AFI") && $comm_person_resp eq "UNASSIGNED UNASSIGNED" && $description eq " ") { $bu_filter = "TRUE";} # no idea why there's a space in description field from GCH instead of empty 
        
              
        if ($bu !~ /^CV-/ && $bu_filter eq "FALSE" && ($geo_person eq "UNASSIGNED UNASSIGNED" || $geo_person eq "" || $geo_person eq "Nathalie Pauly") && $description !~ /^Wait for BU/i &&
            index($opsnames, $comm_person_resp) != -1)
        {
          $followup_gch_list = $followup_gch_list . $gch_number . ";";
          
          $prio = 2;
          $arrayref = $bu eq "NV" ? \@follow_up_nv : \@follow_up;
          
          ##disabled prio sorting for countries
          #if ((index($priocountries, $country) != -1) || $due_date eq "")
          if ($due_date eq "")
          {
            $arrayref = $bu eq "NV" ? \@follow_up_nv_prio : \@follow_up_prio;
            $prio = 1;
          }
          
          
          my %followup_record = (
            gch_number => $gch_number,
            bu => $bu,
            description => $description,
            country => $country,
            due_date => (join '-', reverse split '-', $due_date),
            source => $sheet->Range("R$line")->{Value},
            source_id => $sheet->Range("S$line")->{Value},
            geo_person => $geo_person,
            comm_person_resp => $comm_person_resp,
            count_fu => $sheet->Range("V$line")->{Value},
            count_gfe_successful => $sheet->Range("U$line")->{Value},
            count_gfe_requirement => $sheet->Range("T$line")->{Value},
            prio => $prio,
          );
          
          push @$arrayref,\%followup_record;		
        }
      }
      $line++;
    } until (not $com_id);
  }

  if ($sheetname eq "PE Updates")
  {
    print "Collecting PE Updates\n";
    $line = 5;
    
    ##sort on date at collection
    $sheet->Range("B$line:T$excel_max_row")->Sort({Key1 => $sheet->Range("F$line:F$excel_max_row"), Order1 => xlAscending});
    
    do
    {
      $type = $sheet->Range("D$line")->{Value};
      if ($type)
      {
        $gch_number = $sheet->Range("B$line")->{Value};
        $geo_person = $sheet->Range("P$line")->{Value};
        $description = $sheet->Range("S$line")->{Value};
        $comm_person_resp = $sheet->Range("Q$line")->{Value};
        $country = $sheet->Range("O$line")->{Value};
        $bu = $sheet->Range("H$line")->{Value};
        $prio = (index($priocountries, $country) != -1) ? 1 : 2;
                
        if ($bu !~ /^CV-/ && ($geo_person eq "UNASSIGNED UNASSIGNED" || $geo_person eq "" || $geo_person eq "Nathalie Pauly") && uc($description) ne "BU" &&
            index($opsnames, $comm_person_resp) != -1)
        {
          if (index(lc($description), "merge") != -1)
          {
            $arrayref = \@pe_updates_merged;
          }
          elsif ($bu eq "NV")
          {
            $arrayref = $prio == 1 ? \@pe_updates_nv_prio : \@pe_updates_nv;
          }
          else
          {
            $arrayref = $prio == 1 ? \@pe_updates_prio : \@pe_updates;
          }          
          my %pe_update_record = (
            gch_number => $gch_number,
            bu => $bu,
            description => $description,
            country => $country,
            created_date => (join '-', reverse split '-', $sheet->Range("F$line")->{Value}),
            source => $sheet->Range("I$line")->{Value},
            geo_person => $geo_person,
            prio => $prio,
          );
          push @$arrayref,\%pe_update_record;
          if (!grep($_ == $gch_number, @pe_updatas_gchnumbers)) { push @pe_updatas_gchnumbers, $gch_number; }
        }
      }
      $line++;
    } until (not $type);
    
    print "Creating PE Updates\n";
    $pe_updates_count = 0;
    $pe_updates_nv_count = 0;
    $w = 1;
    
    foreach $type ("PE Update", "PE Update NV")
    {
      $pe_updates_sheet->Range("A$w")->{value} = $type;
      $pe_updates_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
      $w++;
      $pe_updates_sheet->Range("A$w:H$w")->{value} = ["PE#", "Aging", "Task Created Date", "Country", "BU", "Description", "Source", "Geo Person"];
      $pe_updates_sheet->Range("A$w:H$w")->Interior->{ColorIndex} = 36;
      $w++;
      $w_start = $w;
      foreach ($type eq "PE Update" ? (@pe_updates_prio,@pe_updates) : (@pe_updates_nv_prio,@pe_updates_nv))
      {
        my %sr_hash = %{$_};
        $pe_updates_sheet->Range("A$w:H$w")->{value} = [$sr_hash{gch_number}, "=if(c$w,today()-c$w,-9999)", $sr_hash{created_date}, $sr_hash{country}, $sr_hash{bu}, $sr_hash{description}, $sr_hash{source}, $sr_hash{geo_person}];
        $pe_updates_sheet->Range("C$w")->{NumberFormat} = "dd-MMM-yyyy";
        $pe_updates_sheet->Range("B$w")->{NumberFormat} = "#0";
        
        if (index($priocountries, $sr_hash{country}) != -1)
        { 
          $pe_updates_sheet->Range("D$w")->Interior->{ColorIndex} = $prio_color; 
        }
        
        #convert formula to value as workaround for timezone differences in google docs
        $testdate = $pe_updates_sheet->Range("B$w")->{Value};
        $pe_updates_sheet->Range("B$w")->{Value} = $testdate;
        $type eq "PE Update NV" ? $pe_updates_nv_count++ : $pe_updates_count++;
        $w++;
      }
      ### MOVED Sorting to collection
      ###$pe_updates_sheet->Range("A$w_start:H$w")->Sort({Key1 => $pe_updates_sheet->Range("B$w_start:B$w"), Order1 => xlDescending});
      $w++;
    }
       
    print "Collecting updates from mailbox\n";
    
    $pe_updates_sheet->Range("A$w")->{value} = "PE Updates Mailbox";
    $pe_updates_sheet->Columns("A:XX")->{AutoFit}= "True";
    $pe_updates_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
    $w++;
    $pe_updates_sheet->Range("A$w:E$w")->{value} = ["PE#", "Aging", "Date received", "BU", "Subject"];
    $pe_updates_sheet->Range("A$w:E$w")->Interior->{ColorIndex} = 36;
    $w++;
    $w_start = $w;
    $mailbox_update_count = 0;
    $mailbox_nv_update_count = 0;

    $mailbox = seekFolder($outlook->Session, 'RS PP EMEA');
    $inbox = seekFolder($mailbox, 'Inbox');
    $Updates_to_process = seekFolder($inbox, 'Updates');
    $folder = $Updates_to_process->Items;
    for ($i = $Updates_to_process->Items->Count ; $i >=1 ; $i--)
    {
      $item = $folder->Item($i);
      $subject = $item->Subject;
      $sent = $item->CreationTime;
      $sent = join '-', reverse split '-', substr($sent, 0, index($sent, " "));
      $bu = $item->Categories;
      $subject = " $subject";
      $subject =~ s/\D/ /g;
      do
      {
        $gch_number="";
        $extra = 1;
        $pos = index($subject, " 7");
        if (pos == -1) { $pos = index($subject, " 07"); $extra = 2; }
        if ($pos != -1)
        {
          $pos2 = index($subject, " ", $pos + $extra);
          if ($pos2 != -1)
          {
            $gch_number = substr($subject, $pos + $extra, $pos2 - $pos - 1);
          }
          else
          {
            $gch_number = substr($subject, $pos + $extra);
          }
          $subject = substr($subject, $pos + 1);
        }  
      } until ($pos == -1 || length($gch_number) == 9);
      if ($gch_number)
      {
        if (!grep($_ == $gch_number, @mailbox_updatas_gchnumbers)) { push @mailbox_updatas_gchnumbers, $gch_number; }
      }
      $pe_updates_sheet->Range("A$w:E$w")->{value} = [$gch_number, "=today()-c$w", $sent, $bu, $item->Subject];
      $pe_updates_sheet->Range("C$w")->{NumberFormat} = "dd-MMM-yyyy";
      $pe_updates_sheet->Range("B$w")->{NumberFormat} = "#0";
      $bu eq "NV" ? $mailbox_nv_update_count++ : $mailbox_update_count++;
      $w++;
    }
    $w--;
    $pe_updates_sheet->Range("A$w_start:H$w")->Sort({Key1 => $pe_updates_sheet->Range("B$w_start:B$w"), Order1 => xlDescending});
    
    print "Creating Follow Up\n";
        
    $w = 1;
    $followup_count = 0;
    $followup_urgent_count = 0;
    $followup_nv_count = 0;
    $followup_nv_urgent_count = 0;

    foreach $type ("Follow Up", "Follow Up NV")
    {
      $followup_sheet->Range("A$w")->{value} = $type;
      $followup_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
      $w++;
      $followup_sheet->Range("A$w:N$w")->{value} = ["PE#", "Update Locations", "Total Days left", "Due Date", "Country", "BU", "Description", "Source", "Source ID", "Geo Person", "Com Person", "Follow-up Attempt #", "Successful GFE Attempts", "GFE Requirement for Follow Up"];
      $followup_sheet->Range("A$w:N$w")->Interior->{ColorIndex} = 36;
      $w++;
      $w_start = $w;
      foreach ($type eq "Follow Up" ? (@follow_up_prio,@follow_up) : (@follow_up_nv_prio,@follow_up_nv))
      {
        my %sr_hash = %{$_};
        my $update_locations = "";
        if (grep($_ == $sr_hash{gch_number}, @pe_updatas_gchnumbers)) { $update_locations = $update_locations . "; Interface"; }
        if (grep($_ == $sr_hash{gch_number}, @mailbox_updatas_gchnumbers)) { $update_locations = $update_locations . "; Mailbox->Updates"; }
        $update_locations = substr($update_locations,2);
        $followup_sheet->Range("A$w:N$w")->{value} = [$sr_hash{gch_number}, $update_locations, "=if(d$w,d$w-today(),-9999)", $sr_hash{due_date}, $sr_hash{country}, $sr_hash{bu}, $sr_hash{description}, $sr_hash{source}, "\'$sr_hash{source_id}", $sr_hash{geo_person}, $sr_hash{comm_person_resp}, $sr_hash{count_fu}, $sr_hash{count_gfe_successful}, $sr_hash{count_gfe_requirement}];
        $followup_sheet->Range("D$w")->{NumberFormat} = "dd-MMM-yyyy";
        $followup_sheet->Range("C$w")->{NumberFormat} = "#0";
        
         
        #convert formula to value as workaround for timezone differences in google docs
        $testdate = $followup_sheet->Range("C$w")->{Value};
        $followup_sheet->Range("C$w")->{Value} = $testdate;
                
        
        if (index($priocountries, $sr_hash{country}) != -1)
        { 
          $followup_sheet->Range("E$w")->Interior->{ColorIndex} = $prio_color;        
        }
        
        if ($followup_sheet->Range("C$w")->{value} <= 0)
        {
          ### Disabled aging coloring due to new way of working
          #$followup_sheet->Range("C$w")->Interior->{ColorIndex} = $prio_color;
          $sr_hash{bu} eq "NV" ? $followup_nv_urgent_count++ : $followup_urgent_count++;          
        }
        
        if ($followup_sheet->Range("C$w")->{value} == -9999)
        {
          $followup_sheet->Range("C$w")->Interior->{ColorIndex} = $prio_color;
          $followup_sheet->Range("D$w")->Interior->{ColorIndex} = $prio_color;
          $followup_sheet->Range("D$w")->{value} = "Due date not set!";
          $followup_sheet->Range("D$w")->Font->{ColorIndex} = $alert_color;          
        }
        
        
        if (grep($_ == $sr_hash{gch_number}, (@pe_updatas_gchnumbers,@mailbox_updatas_gchnumbers))) { $followup_sheet->Range("A$w")->Font->{ColorIndex} = $update_color; }
        $sr_hash{bu} eq "NV" ? $followup_nv_count++ : $followup_count++;
        $w++;
      }
      
      ###moved sorting to collection to keep prio order
      ###$followup_sheet->Range("A$w_start:N$w")->Sort({Key1 => $followup_sheet->Range("C$w_start:C$w")});
      $w++;
    }
    $followup_sheet->Columns("A:XX")->{AutoFit}= "True";
  }
  
   
  ### Completions

  if ($sheetname eq "Completions" && $do_completions eq "TRUE")
  {
    $completions_count = 0;
    $completions_urgent_count = 0;
    $completions_nv_count = 0;
    $completions_nv_urgent_count = 0;
    $line = 2;
       
    @completions = ();
    @full_completions = ();
    $arrayref = \@completions;
    print "Collecting completions\n";
    while ($gch_number = $sheet->Range("A$line")->{Value}) 
    {
      $completion_type = 0;
      $gch_number = substr $gch_number, 1, 9;
      $pe_task_type = $sheet->Range("Z$line")->{Value};
      $bu = $sheet->Range("G$line")->{Value};
      $geo_person = $sheet->Range("I$line")->{Value};
      $investigation_status = $sheet->Range("U$line")->{Value};
      $pe_task_status = $sheet->Range("AA$line")->{Value};
      $pe_aging = $sheet->Range("F$line")->{Value};
      
      print "Processing: ".$gch_number."\n";
                      
        if ($pe_aging >= ($bu eq "CRDM" ? 90 : $bu eq "Xomed" ? 55 : 60))
        {
          $prio = 1;
        }
        else 
        {
          $prio = 2; 
        }
        
        $prod_returned_collect = "";
        $rational_no_return_collect = "";
        $open_task_collect = "";
        $open_comm_collect = "";
        ### Collect product returned and rationals + open tasks, note that as a side effect this removes duplicate PE's,
        ### if more information is needed from those duplicate lines add it to this loop
        while (index ($sheet->Range("A$line")->{Value},$gch_number) != -1)
        {
          $prod_returned = $sheet->Range("O$line")->{Value};
          if ($prod_returned && $prod_returned_collect !~ /$prod_returned/)
          {
            $prod_returned_collect = $prod_returned_collect ? "$prod_returned_collect, $prod_returned" : $prod_returned;
          }	
          $rational_no_return = $sheet->Range("P$line")->{Value};
          if ($rational_no_return && $rational_no_return_collect !~ /$rational_no_return/ )
          {
            $rational_no_return_collect = $rational_no_return_collect ? "$rational_no_return_collect, $rational_no_return" : $rational_no_return;
          }
          $open_task = $sheet->Range("Z$line")->{Value};
          $open_task_status = $sheet->Range("AA$line")->{Value};
          if ($open_task && $open_task ne "Geo Event Review" && $open_task_status ne "Complete" && $open_task_collect !~ /$open_task/ )
          {
            $open_task_collect = $open_task_collect ? "$open_task_collect; $open_task" : $open_task;
            if ($prio < 3) { $prio = $prio + 2;}
          }
          $open_comm = $sheet->Range("AB$line")->{Value};
          $open_comm_status = $sheet->Range("AC$line")->{Value};
          if ($open_comm && $open_comm_status ne "Complete" && $open_comm_status ne "Not Required") 
          {
            $open_comm_pers_resp = $sheet->Range("AD$line")->{Value};
            $open_comm_item = $open_comm . " (" . $open_comm_pers_resp . ")";
            if (index($open_comm_collect,$open_comm_item) == -1 )
            {
              $open_comm_collect = $open_comm_collect ? "$open_comm_collect; $open_comm_item" : $open_comm_item;
              if ($prio < 3) { $prio = $prio + 2;}
            }
          }
          if ($open_task eq "Geo Event Review") 
          {
            if ($geo_person eq "UNASSIGNED UNASSIGNED" || $geo_person eq "" || $geo_person eq "Nathalie Pauly")
            {
              $completion_type = 2;
              if ($investigation_status eq "Complete")
              {
                if ($open_task_status eq "New" || $open_task_status eq "Re-Open" || $open_task_status eq "In Progress")
                {
                  $completion_type = 1;
                }
              }
            }
          }
          $line++;
        }
        
      $line--;
        
      
        
      my %completion = (
        gch_number => $gch_number,
        bu => $bu,
        geo_person => $geo_person,
        pe_days_open => $sheet->Range("D$line")->{Value},
        pe_ageing => $pe_aging,
        country => $sheet->Range("J$line")->{Value},
        product_returned => $prod_returned_collect,
        rational_no_return => $rational_no_return_collect,
        source_system => $sheet->Range("W$line")->{Value},
        prio => $prio,
        open_tasks => $open_task_collect,
        open_comms => $open_comm_collect,
        investigation_status => $investigation_status,
        svc_order_tech_completion_date => $sheet->Range("X$line")->{Value},
        notification_complete_date => $sheet->Range("Y$line")->{Value},
      );
         
      if ($completion_type == 1) 
      { 
        push @$arrayref,\%completion; 
        if ($prio < 3) {$bu eq "NV" ? $completions_nv_count++ : $completions_count++;}
        if ($prio == 1) {$bu eq "NV" ? $completions_nv_urgent_count++ : $completions_urgent_count++ };
      }
     # if ($completion_type > 0) {
        push @full_completions,\%completion;                         
    #  }

      $line++;
    }
    
        
    ## create completions sheet
    @buList = ("Xomed","PSS","MAE","Spinal","CRDM","Cryocath","AFI","NEUROMOD","NeuroSurgery","NV");
    $w = 1;
    
        
    ## TODO: Make sure completions array is sorted by BU so only one iteration is needed
    foreach my $current_bu (@buList) {
      $completions_sheet->Range("A$w")->{value} = $current_bu;
      $completions_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
      $w++;
      $completions_sheet->Range("A$w:G$w")->{value} = ["PE#", "Total Days open", "PE aging", "Country", "Product Returned", "Rational Not Returned", "Source"];
      $completions_sheet->Range("A$w:G$w")->Interior->{ColorIndex} = 36;
      $w++;
         
      foreach (@completions) {
        my %cp_hash = %{$_};        
        if (index($followup_gch_list,$cp_hash{gch_number}) == -1)
        {
          if ($cp_hash{bu} eq $current_bu)
          {
            if ($cp_hash{prio} < 3) {
              $completions_sheet->Range("A$w:G$w")->{value} = [$cp_hash{gch_number}, $cp_hash{pe_days_open}, $cp_hash{pe_ageing}, $cp_hash{country}, $cp_hash{product_returned}, $cp_hash{rational_no_return}, $cp_hash{source_system}];
              if ($cp_hash{prio} == 1) { $completions_sheet->Range("C$w")->Interior->{ColorIndex} = $prio_color; }
              $w++;
            }             
          }
        }
      }
      $w++;
    }
    
    $x = 1;
    $qc_completions_sheet->Range("A$x")->{value} = "Quality Check - Completions";
    $qc_completions_sheet->Range("A$x")->Interior->{ColorIndex} = 35;
    $x++;
    $qc_completions_sheet->Range("A$x:N$x")->{value} = ["PE#", "Total Days open", "PE aging", "Country", "BU", "Product Returned", "Rational Not Returned", "Open Comms", "Open Tasks", "Source", "Investigation Status", "Svc Order Technical Completion Date-S&R", "Notification Complete Date-S&R", "GEO Person Responsible"];
    $qc_completions_sheet->Range("A$x:N$x")->Interior->{ColorIndex} = 36;
    $x++;
    
    foreach (@full_completions) 
      {
        my %cp_hash = %{$_};
        $qc_completions_sheet->Range("A$x:N$x")->{value} = [$cp_hash{gch_number}, $cp_hash{pe_days_open}, $cp_hash{pe_ageing}, $cp_hash{country}, $cp_hash{bu}, $cp_hash{product_returned}, $cp_hash{rational_no_return}, $cp_hash{open_comms}, $cp_hash{open_tasks}, $cp_hash{source_system}, $cp_hash{investigation_status}, $cp_hash{svc_order_tech_completion_date}, $cp_hash{notification_complete_date}, $cp_hash{geo_person}];
        $qc_completions_sheet->Range("L$x")->{NumberFormat} = "dd-MMM-yyyy";
        $qc_completions_sheet->Range("M$x")->{NumberFormat} = "dd-MMM-yyyy";
        if ($cp_hash{prio} % 2 > 0) { $qc_completions_sheet->Range("C$x")->Interior->{ColorIndex} = $prio_color; }
        $x++;       
      }
    
    $completions_sheet->Columns("A:XX")->{AutoFit}= "True";    
    $qc_completions_sheet->Columns("A:XX")->{AutoFit}= "True";        
  }
  
  @intake_others = ();
  #@intake_neuro = ();
  @intake_nv = ();
    
  if ($sheetname eq "Intakes")
  {
    print "Collecting intakes\n";
    $line = 5;
    while ($gch_number = $sheet->Range("B$line")->{Value})
    {
      last if length($gch_number) < 9;
      $geo_person = $sheet->Range("I$line")->{Value};
      $task_status = $sheet->Range("L$line")->{Value};
      
      if ($task_status eq "New" && ($geo_person eq "UNASSIGNED UNASSIGNED" || $geo_person eq "Nathalie Pauly" || $geo_person eq ""))
      { 
        $bu = $sheet->Range("F$line")->{Value};
        $country = $sheet->Range("G$line")->{Value};
        $patient_status = $sheet->Range("V$line")->{Value};
        
        if ($patient_status ne "" && $patient_status ne "Alive" || index($priocountries, $country) != -1) { $prio = 1; }
        else { $prio = 2; }
        
        #if ($bu eq "NEUROMOD") { $arrayref = \@intake_neuro; }
        if ($bu eq "NV") { $arrayref = \@intake_nv; }
        else { $arrayref = \@intake_others; }
        
        my %intake_record = (
          gch_number => $gch_number,
          bu => $bu,
          notified_date => (join '-', reverse split '-', $sheet->Range("E$line")->{Value}),
          country => $country,
          description => $sheet->Range("H$line")->{Value},
          source => $sheet->Range("Q$line")->{Value},
          source_id => $sheet->Range("R$line")->{Value},
          patient_status => $patient_status,
          prio => $prio,
          geo_person => $geo_person,
        );
        push @$arrayref,\%intake_record;
      }
      $line++;
    }
    
    print "Creating Intakes\n";
      
    $w = 1;
    $intakes_count = 0;
    $intakes_urgent_count = 0;
    $intakes_nv_count = 0;
    $intakes_nv_urgent_count = 0;
    $intakes_sr_count = 0;
    $intakes_sr_urgent_count = 0;
    
    #foreach $type ("Intakes", "Intakes Neuro", "Intakes NV")
    foreach $type ("Intakes", "Intakes NV")
    {
      $intakes_sheet->Range("A$w")->{value} = $type;
      $intakes_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
      $w++;
      $intakes_sheet->Range("A$w:J$w")->{value} = ["PE#", "Aging", "Notified Date", "BU", "Country", "Patient Status", "Description", "Source", "Source ID", "Geo Person Responsible (PE)"];
      $intakes_sheet->Range("A$w:J$w")->Interior->{ColorIndex} = 36;
      $w++;
      $w_start = $w;
      
      if ($type eq "Intakes") { $arrayref = \@intake_others; }
      #elsif ($type eq "Intakes Neuro") { $arrayref = \@intake_neuro; }
      elsif ($type eq "Intakes NV") { $arrayref = \@intake_nv; }
      
      ### TODO: Replace this loop with proper sorting in the arrays
      foreach $priority (1,2) {
        foreach (@$arrayref)
        {
          my %sr_hash = %{$_};
          $urgent = $sr_hash{prio};
          if ($priority == $urgent) {
            $intakes_sheet->Range("A$w:J$w")->{value} = [$sr_hash{gch_number}, "=today()-c$w", $sr_hash{notified_date}, $sr_hash{bu}, $sr_hash{country}, $sr_hash{patient_status}, $sr_hash{description}, $sr_hash{source}, "\'$sr_hash{source_id}", $sr_hash{geo_person}];
            $intakes_sheet->Range("C$w")->{NumberFormat} = "dd-MMM-yyyy";
            $intakes_sheet->Range("B$w")->{NumberFormat} = "####";
            #convert formula to value as workaround for timezone differences in google docs
            $testdate = $intakes_sheet->Range("B$w")->{Value};
            $intakes_sheet->Range("B$w")->{Value} = $testdate;
                        
            if ($weekday eq "Mon" || $weekday eq "Tue" ) { $compareDays = 4 } else { $compareDays = 2 }
            
            if ($intakes_sheet->Range("B$w:B$w")->{value} > $compareDays)
            {
              #disabled aging coloring due to new way of working
              #$intakes_sheet->Range("B$w")->Interior->{ColorIndex} = $prio_color;
              $urgent = 1;
            }
            
            if ($sr_hash{patient_status} ne "" && $sr_hash{patient_status} ne "Alive") 
            { 
              $intakes_sheet->Range("A$w:B$w,F$w")->Interior->{ColorIndex} = $prio_color;
              $urgent = 1;
            }
            
            elsif (index($priocountries, $sr_hash{country}) != -1)
            { 
              $intakes_sheet->Range("E$w")->Interior->{ColorIndex} = $prio_color;
              $urgent = 1;
            }
                
            if ($sr_hash{bu} eq "NV") {$intakes_nv_count++; }
            elsif ($sr_hash{source} eq "SAP ECC Service and Repair") { $intakes_sr_count++; }
            else {$intakes_count++; }
            
            if ($urgent == 1) {
              if ($sr_hash{bu} eq "NV") { $intakes_nv_urgent_count++; }
              elsif ($sr_hash{source} eq "SAP ECC Service and Repair") { $intakes_sr_urgent_count++; }
              else {$intakes_urgent_count++; }
            }            
            $w++;
          }
        }
      }
      $w++;      
    }
    $intakes_sheet->Columns("A:XX")->{AutoFit}= "True";
  }
  
  ### Staging Records Communications
  if ($sheetname eq "SR Communications")
  {
    print "Collecting SR Communications...\n";
    $sr_comms_count = 0;
    $sr_comms_urgent_count = 0;
    $line = 5;
    
    @sr_comms = ();
    $arrayref = \@sr_comms;
    
    while ($gch_number = $sheet->Range("B$line")->{Value})
    {
      $gch_number = substr $gch_number, 1, 9;
      $comm_person_resp = $sheet->Range("H$line")->{Value};
      
      $match_ops = index($opsnames, $comm_person_resp) != -1 ? 1 : 0;
      
      #print "Processing SR Comm GCH Number:$gch_number, Comm Pers resp: #$comm_person_resp#, Match: $match_ops \n";
      
      if (index($opsnames, $comm_person_resp) != -1)
      {
        my %sr_comm_record = (
            gch_number => $gch_number,
            due_date => (join '-', reverse split '-', $sheet->Range("M$line")->{Value}),
            bu => $sheet->Range("C$line")->{Value},
            description => $sheet->Range("I$line")->{Value},
            general_text => $sheet->Range("N$line")->{Value},
            comm_person_resp => $comm_person_resp,
            count_fu => $sheet->Range("J$line")->{Value},
            count_total_fu => $sheet->Range("K$line")->{Value},            
          );
          
          push @$arrayref,\%sr_comm_record;	
          $sr_comms_count++;
      }
      $line++;
    }
      
      $w = 1;
      $sr_comms_sheet->Range("A$w")->{value} = $sheetname;
      $sr_comms_sheet->Range("A$w")->Interior->{ColorIndex} = 35;
      $w++;
      $sr_comms_sheet->Range("A$w:I$w")->{value} = ["SR#", "Total Days left", "Due Date", "BU", "Description", "General Text", "Comm Person", "Follow-up Attempt #", "Total # of Attempts"];
      $sr_comms_sheet->Range("A$w:I$w")->Interior->{ColorIndex} = 36;
      $w++;
      $w_start = $w;
      
      foreach (@sr_comms)
      {
        my %sr_hash = %{$_};
        $sr_comms_sheet->Range("A$w:I$w")->{value} = [$sr_hash{gch_number}, "=if(c$w,c$w-today(),-9999)", $sr_hash{due_date}, $sr_hash{bu}, $sr_hash{description}, $sr_hash{general_text}, $sr_hash{comm_person_resp}, $sr_hash{count_fu}, $sr_hash{count_fu_total}];
        $sr_comms_sheet->Range("C$w")->{NumberFormat} = "dd-MMM-yyyy";
        $sr_comms_sheet->Range("B$w")->{NumberFormat} = "#0";
        #convert formula to value as workaround for timezone differences in google docs
        $testdate = $sr_comms_sheet->Range("B$w")->{Value};
        $sr_comms_sheet->Range("B$w")->{Value} = $testdate;
        if ($sr_comms_sheet->Range("B$w")->{value} <= 0)
        {
          $sr_comms_urgent_count++;
          #disabled aging coloring due to new way of working
          #$sr_comms_sheet->Range("B$w")->Interior->{ColorIndex} = $prio_color;
        }
        $w++;
      }
      $sr_comms_sheet->Range("A$w_start:I$w")->Sort({Key1 => $sr_comms_sheet->Range("B$w_start:B$w")});
      $sr_comms_sheet->Columns("A:XX")->{AutoFit}= "True";     
  }
  
  if ($sheetname eq "Cover page")
  { 
    $crdm_quickhits = $sheet->Range("C10")->{Value} + $sheet->Range("H10")->{Value};
  }
}

## Add updates about merged cases to Staging Records sheet
$sr_sheet_line++;
$sr_sheet->Range("A$sr_sheet_line")->{value} = "Merged by BU";
$sr_sheet->Range("A$sr_sheet_line")->Interior->{ColorIndex} = 35;
$sr_sheet_line++;
$sr_sheet->Range("A$sr_sheet_line:H$sr_sheet_line")->{value} = ["PE#", "Aging", "Task Created Date", "Country", "BU", "Description", "Source", "Geo Person"];
$sr_sheet->Range("A$sr_sheet_line:H$sr_sheet_line")->Interior->{ColorIndex} = 36;
$sr_sheet_line++;

foreach (@pe_updates_merged)
{
  my %sr_hash = %{$_};
  $sr_sheet->Range("A$sr_sheet_line:H$sr_sheet_line")->{value} = [$sr_hash{gch_number}, "=if(c$sr_sheet_line,today()-c$sr_sheet_line,-9999)", $sr_hash{created_date}, $sr_hash{country}, $sr_hash{bu}, $sr_hash{description}, $sr_hash{source}, $sr_hash{geo_person}];
  $sr_sheet->Range("C$sr_sheet_line")->{NumberFormat} = "dd-MMM-yyyy";
  $sr_sheet->Range("B$sr_sheet_line")->{NumberFormat} = "#0";
    
  #convert formula to value as workaround for timezone differences in google docs
  $testdate = $sr_sheet->Range("B$sr_sheet_line")->{Value};
  $sr_sheet->Range("B$sr_sheet_line")->{Value} = $testdate;
  $sr_sheet_line++;
}

### Layout update
$sr_sheet->Columns("A:XX")->{AutoFit}= "True";    

print "Creating Coversheet\n";
$coversheet->Range("B1:C1")->{value} = ["Total Count","Urgent Count"];
$line = 2;
$coversheet->Range("A$line:C$line")->{value} = ["Staging Records", $staging_record_count, $staging_record_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Staging Records - Service & Repair", $staging_record_sr_count, $staging_record_sr_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Intakes", $intakes_count, $intakes_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Intakes - Service & Repair", $intakes_sr_count, $intakes_sr_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Intakes NV", $intakes_nv_count, $intakes_nv_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["PIRs to process (Mailbox SR Promotion)", $pirs_to_process_count, $pirs_to_process_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Completions", $completions_count, $completions_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Completions NV", $completions_nv_count, $completions_nv_urgent_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Completions CRDM Quickhits", $crdm_quickhits, $crdm_quickhits]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["PE Updates", $pe_updates_count, $pe_updates_count]; $line++; 
$coversheet->Range("A$line:C$line")->{value} = ["PE Updates NV", $pe_updates_nv_count, $pe_updates_nv_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Mailbox Updates", $mailbox_update_count, $mailbox_update_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Mailbox Updates NV", $mailbox_nv_update_count, $mailbox_nv_update_count]; $line++;
$coversheet->Range("A$line:C$line")->{value} = ["Follow Up", $followup_count, $followup_urgent_count]; $line++; 
$coversheet->Range("A$line:C$line")->{value} = ["Follow Up NV", $followup_nv_count, $followup_nv_urgent_count]; $line++; 
$coversheet->Range("A$line:C$line")->{value} = ["Staging Record Communications Count", $sr_comms_count, $sr_comms_urgent_count]; $line++; 
$coversheet->Columns("A:XX")->{AutoFit}= "True";

$coversheet->Activate();

$Book->Quit();

#Ignore popup when file already exists
$Excel->{DisplayAlerts} = 0;

$resultFilename = $script_location."opsman\\wip_ops_team$testing\_GCH6_$current_year$\-$current_month\-$current_day\_$current_hour\-$current_min\-$current_sec.xlsx";

print "Saving: ",$resultFilename,"\n";
$WipBook->SaveAs($resultFilename);
$WipBook->Quit();
#$Excel->Quit();

#print "Emailing the wipman report\n";
#$message = $outlook->CreateItem(0);
#$message->{'To'} = 'Reijntjens, Guido;Haagen, Tom;Kokhuis,Tom;Daemen, Raoul;Heemskerk, Ernest;Maurer, Stijn;Weijzen, Sylvia;van Diesen, Wolf;Meukens, Jeroen;Augenbroe, Ralf;Severens, Natascha;Verboeket, Yana;van der Ven, Wesley;Smeekes, Conny';
#$message->{'Subject'} = "Wipman $current_year-$current_month-$current_day";
#$message->{'Body'} = 'Attached today\'s Wipman';
#$attachments = $message->Attachments();
#$attachments->Add($resultFilename);
#$message->Send();

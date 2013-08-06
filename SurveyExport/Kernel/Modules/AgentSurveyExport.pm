# --
# Kernel/Modules/AgentSurveyExport.pm - a survey export module
# Copyright (C) 2013 tuxwerk OHG, http://tuxwerk.de/
# --
# This software comes with ABSOLUTELY NO WARRANTY. For details, see
# the enclosed file COPYING for license information (AGPL). If you
# did not receive this file, see http://www.gnu.org/licenses/agpl.txt.
# --

package Kernel::Modules::AgentSurveyExport;

use strict;
use warnings;

use Kernel::System::Survey;
use Kernel::System::CSV;
use Kernel::System::HTMLUtils;

use vars qw($VERSION);
$VERSION = qw($Revision: 1.13 $) [1];

sub new {
    my ( $Type, %Param ) = @_;

    # allocate new hash for object
    my $Self = {};
    bless( $Self, $Type );

    # get common objects
    %{$Self} = %Param;

    # check needed objects
    for my $Object (qw(ParamObject TicketObject DBObject LayoutObject LogObject ConfigObject)) {
        if ( !$Self->{$Object} ) {
            $Self->{LayoutObject}->FatalError( Message => "Got no $Object!" );
        }
    }
    $Self->{SurveyObject} = Kernel::System::Survey->new(%Param);
    $Self->{CSVObject}    = Kernel::System::CSV->new(%Param);
    $Self->{HTMLUtilsObject} = Kernel::System::HTMLUtils->new(%Param);

    return $Self;
}

sub Run {
    my ( $Self, %Param ) = @_;

    my $SurveyID = $Self->{ParamObject}->GetParam( Param => "SurveyID" );

    my @Questions = $Self->{SurveyObject}->QuestionList(
        SurveyID => $SurveyID
    );
    my @CSVHead;
    my @CSVData;

    push @CSVHead, 'Send time';
    push @CSVHead, 'Vote time';
    push @CSVHead, 'Queue';
    push @CSVHead, 'Ticket';
    push @CSVHead, 'Ticket owner';
    foreach(@Questions) {
	push @CSVHead, $_->{Question};
    }

    my @List = $Self->{SurveyObject}->VoteList( SurveyID => $SurveyID );
    # Sendezeit/SendTime, Abstimmungszeit/VoteTime
    for my $Vote (@List) {
	my @Data;
	my %Ticket = $Self->{TicketObject}->TicketGet( TicketID => $Vote->{TicketID} );
	push @Data, $Vote->{SendTime};
	push @Data, $Vote->{VoteTime};
	push @Data, $Ticket{Queue};
	push @Data, "#".$Ticket{TicketNumber};
	push @Data, $Ticket{Owner};

	for my $Question (@Questions) {
	    my $Answer = "";
	    my @Answers = $Self->{SurveyObject}->VoteGet(
		RequestID  => $Vote->{RequestID},
		QuestionID => $Question->{QuestionID},
            );
	    if ( $Question->{Type} eq 'Radio' || $Question->{Type} eq 'Checkbox' ) {
		for my $Row (@Answers) {
                    my %AnswerText = $Self->{SurveyObject}->AnswerGet( AnswerID => $Row->{VoteValue} );
                    $Answer .= $AnswerText{Answer};
                }
	    }
	    elsif ( $Question->{Type} eq 'YesNo' || $Question->{Type} eq 'Textarea' ) {
		$Answer = $Answers[0]->{VoteValue};
		# clean html
                if ( $Question->{Type} eq 'Textarea' && $Answer ) {
                    $Answer =~ s{\A\$html\/text\$\s(.*)}{$1}xms;
                    $Answer = $Self->{HTMLUtilsObject}->ToAscii( String => $Answer );
		    # make excel linebreak in cell work
		    $Answer =~ s/\n/\r/g; 
                }
	    }
	    push @Data, $Answer;
	}

	push @CSVData, \@Data;
    }

    # get Separator from language file
    my $UserCSVSeparator = $Self->{LayoutObject}->{LanguageObject}->{Separator};

    if ( $Self->{ConfigObject}->Get('PreferencesGroups')->{CSVSeparator}->{Active} ) {
        my %UserData = $Self->{UserObject}->GetUserData( UserID => $Self->{UserID} );
        $UserCSVSeparator = $UserData{UserCSVSeparator};
    }

    my $CSV = $Self->{CSVObject}->Array2CSV(
	Head      => \@CSVHead,
	Data      => \@CSVData,
	Separator => $UserCSVSeparator,
    );

    return $Self->{LayoutObject}->Attachment(
	Filename    => "survey.csv",
	ContentType => "text/csv; charset=utf8",
	Content     => $CSV,
    );
}

1;

// =DOC FILE Interface for the DArea class
// =DOC TEXT Declares DArea class to monitor the status of an area during fumigation
//////////////////////////////////////////////////////////////////

// Area.h: interface for the DArea class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_AREA_H__B082E2E9_9ED9_4C84_89EA_248861072EEA__INCLUDED_)
#define AFX_AREA_H__B082E2E9_9ED9_4C84_89EA_248861072EEA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "AreaData.h"
#include "AreaMonPtData.h"


// =DOC SECTION Class DArea
// =DOC SECTION Determine what the area status is with the current data
// =DOC TEXT class DArea extends nothing
// =DOC TEXT Determine what the area status is with the current data
class DArea  
{
public:
	// =DOC SECTION Method GetAreaStatus( LPAREA_DATA lpAreaData ): BOOL
	// =DOC SECTION Get the area status for the current data
	// =DOC TEXT public  GetAreaStatus( LPAREA_DATA lpAreaData ): BOOL
	// =DOC TEXT LPAREA_DATA lpAreaData: A carrier structure for the area status
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Get the area status for the current data
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
	BOOL GetAreaStatus( LPAREA_DATA lpAreaData );

   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DArea();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
	virtual ~DArea();

protected:
   // =DOC SECTION Method AddProfume(AREA_MON_PT_DATA_ARRAY* padataAreaMonPt): BOOL
   // =DOC SECTION Check if monitoring points for an area are within 30 minutes of the
   // =DOC SECTION longest elapsed time
   //////////////////////////////////////////////////////////////////
	BOOL CheckElapsedTimes(AREA_MON_PT_DATA_ARRAY* padataAreaMonPt);

   // =DOC SECTION Method AddProfume( const double &dVolume, const double &dFutureTargetConcentration, const double &dFutureConcentration ): double
   // =DOC SECTION Determine the amount of ProFume to add to reach the target concentration
   // =DOC TEXT protected AddProfume( const double &dVolume, const double &dFutureTargetConcentration, const double &dFutureConcentration ): double
   // =DOC TEXT const double &dVolume: Volume of the area (m^3)
   // =DOC TEXT const double &dFutureTargetConcentration: Target concentration for 1 hour in the future
   // =DOC TEXT const double &dFutureConcentration: Expected concentration for 1 hour in the future, given current conditions
   // =DOC TEXT Returns double: Profume to add (kg)
   // =DOC TEXT Determine the amount of ProFume to add to reach the target concentration
   //////////////////////////////////////////////////////////////////
   double AddProfume( const double &dVolume,
                      const double &dFutureTargetConcentration,
                      const double &dFutureConcentration );

   // =DOC SECTION Method GetMeanValues( double &dAchievedConcentrationTime, double &dAchievedDose, double &dHlt, double &dProjectedConcentrationTime, double &dProjectedTime, double &dFutureTargetConcentration, double &dFutureConcentration, double &dElapsedTime, AREA_MON_PT_DATA_ARRAY* padataAreaMonPt ): void
   // =DOC SECTION Calculate the average data values for the area from its monitor points
   // =DOC TEXT protected GetMeanValues( double &dAchievedConcentrationTime, double &dAchievedDose, double &dHlt, double &dProjectedConcentrationTime, double &dProjectedTime, double &dFutureTargetConcentration, double &dFutureConcentration, double &dElapsedTime, AREA_MON_PT_DATA_ARRAY* padataAreaMonPt ): void
   // =DOC TEXT double &dAchievedConcentrationTime: Achieved CT
   // =DOC TEXT double &dAchievedDose: Achieved dose
   // =DOC TEXT double &dHlt: Half-loss time for the area
   // =DOC TEXT double &dProjectedConcentrationTime: Projected CT
   // =DOC TEXT double &dProjectedTime: Projected time
   // =DOC TEXT double &dFutureTargetConcentration: Target concentration for 1 hour in the future
   // =DOC TEXT double &dFutureConcentration: Expected concentration for 1 hour in the future, given current conditions
   // =DOC TEXT double &dElapsedTime: Elapsed time since start of exposure (h)
   // =DOC TEXT AREA_MON_PT_DATA_ARRAY* padataAreaMonPt: Carrier structure for the Area's Monitor Point data values
   // =DOC TEXT Returns nothing
   // =DOC TEXT Calculate the average data values for the area from its monitor points
   //////////////////////////////////////////////////////////////////
	void GetMeanValues( double &dAchievedConcentrationTime,
                       double &dAchievedDose,
                       double &dHlt,
                       double &dProjectedConcentrationTime,
                       double &dProjectedTime,
                       double &dFutureTargetConcentration,
                       double &dFutureConcentration,
                       double &dElapsedTime,
                       AREA_MON_PT_DATA_ARRAY* padataAreaMonPt );
private:
};


#endif // !defined(AFX_AREA_H__B082E2E9_9ED9_4C84_89EA_248861072EEA__INCLUDED_)

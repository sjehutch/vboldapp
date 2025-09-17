// =DOC FILE Interface for the DMonitorPoint class
// =DOC TEXT Declares DMonitorPoint class to determine the amount of profume required for a fumigation
//////////////////////////////////////////////////////////////////

// MonitorPoint.h: interface for the DMonitorPoint class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MONITORPOINT_H__572A8864_59A5_4D4E_8411_6CBC31244F93__INCLUDED_)
#define AFX_MONITORPOINT_H__572A8864_59A5_4D4E_8411_6CBC31244F93__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



#include "DasPestData.h"
#include "MonitorPoint2PointData.h"
#include "MonitorPointData.h"
#include "PestData.h"

enum PEAK_PATTERN;


// =DOC SECTION Class DMonitorPoint
// =DOC SECTION Determine what the point to point hlt and overall monitor point status is with the current data
// =DOC TEXT class DMonitorPoint extends nothing
// =DOC TEXT Determine what the point to point hlt and overall monitor point status is with the current data
//////////////////////////////////////////////////////////////////
class DMonitorPoint  
{
public:
	// =DOC SECTION Method GetPt2PtHlt( const LPMONITOR_POINT_2_POINT_DATA lpdataMonPt2Pt ): BOOL
	// =DOC SECTION Get the point to point HLT for the current data
	// =DOC TEXT public  GetPt2PtHlt( const LPMONITOR_POINT_2_POINT_DATA lpdataMonPt2Pt ): BOOL
	// =DOC TEXT LPMONITOR_POINT_2_POINT_DATA lpdataMonPt2Pt: A carrier structure for the point to point HLT
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Get the point to point HLT for the current data
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
	BOOL GetPt2PtHlt( const LPMONITOR_POINT_2_POINT_DATA lpdataMonPt2Pt );


	// =DOC SECTION Method GetMonitorPointStatus( const LPMONITOR_POINT_DATA lpdataMonPt ): BOOL
	// =DOC SECTION Get the monitor point status for the current data
	// =DOC TEXT public  GetMonitorPointStatus( const LPMONITOR_POINT_DATA lpdataMonPt ): BOOL
	// =DOC TEXT LPMONITOR_POINT_DATA lpdataMonPt: A carrier structure for the monitor point status
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Get the monitor point status for the current data
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
	// Jett 6/5/03 modified declaration to receive two new parameters
	BOOL GetMonitorPointStatus( const LPMONITOR_POINT_DATA lpdataMonPt, const double dInitConc, const double dOHLT );


   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
   DMonitorPoint();
	

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DMonitorPoint();

protected:
	// Jett 6/10/03 added new function to get our TargetDose using formula similar to Part I Step 10C
	double GetTargetDose(const double dInitialConcentration, const double dnValue, const double dOHLT, const double dExposureTime);

   // =DOC SECTION Struct CalcInfo 
   // =DOC SECTION Carrier structure containing calculation information for a monitor point
   // =DOC TEXT struct CalcInfo extends nothing
   // =DOC TEXT Carrier structure containing calculation information for a monitor point
   //////////////////////////////////////////////////////////////////
   struct CalcInfo
   {
      double dAchievedCT;
      double dAchievedDose;
      double dActualOHLT;
      double dActualEHLT;
      double dRemainingCT;
      double dProjectedCT;
      double dDeltaDose;
      double dProjectedFumigationTime;
      double dFutureC;
      double dFutureDose;
      double dRemainingDose;
      double dNewTargetC;
	  //JNM 04/16/03 - Fum. Enh. 2003 - Added this new variable to hold the n value
	  //	of the current pest.
	  double dPestNValue;
	  //Jett 6/10/03 - added dInitConc so we can use the new equation to get 
	  //CnTTargetD
      double dInitConc;
      double dPlannedHLT;
   };


	// =DOC SECTION Method DoMonitorPointCalculations( const LPMONITOR_POINT_DATA lpdataMonPt, CalcInfo* pCalcInfo ): BOOL
	// =DOC SECTION Perform the Part II calculations for a monitor point
	// =DOC TEXT public DoMonitorPointCalculations( const LPMONITOR_POINT_DATA lpdataMonPt, CalcInfo* pCalcInfo ): BOOL
	// =DOC TEXT const LPMONITOR_POINT_DATA lpdataMonPt: A carrier structure for the data for a monitor point
	// =DOC TEXT CalcInfo* pCalcInfo: A carrier structure for the calculation results for a monitor point
   // =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Perform the Part II calculations for a monitor point
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
   BOOL DoMonitorPointCalculations(const LPMONITOR_POINT_DATA lpdataMonPt, CalcInfo* pCalcInfo);


   // =DOC SECTION Method GetCTAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc ): double
   // =DOC SECTION Calculate the dose achieved using the current data
   // =DOC TEXT protected  GetCTAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc ): double
   // =DOC TEXT CONCENTRATION_DATA_ARRAY *padataConc: Array of concentration values
   // =DOC TEXT Returns double: Achieved dose (g/m^3)
   // =DOC TEXT Calculate the CT achieved using the current data.
   // =DOC TEXT The algorithm is following Part II, step 1 equations, as of 2001-07-23.
   //////////////////////////////////////////////////////////////////
   double GetCTAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc );


   // =DOC SECTION Method GetDoseAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc ): double
   // =DOC SECTION Calculate the dose achieved using the current data with concentration raised to n
   // =DOC TEXT protected  GetDoseAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc ): double
   // =DOC TEXT CONCENTRATION_DATA_ARRAY *padataConc: Array of concentration values
   // =DOC TEXT Returns double: Achieved N dose (g/m^3)
   // =DOC TEXT Calculate the dose achieved using the current data.
   // =DOC TEXT The algorithm is following Part II, step 1 equations, as of 2001-07-23.
   //////////////////////////////////////////////////////////////////
   double GetDoseAchieved(const CONCENTRATION_DATA_ARRAY *pdataConc);
	

   // =DOC SECTION Method GetActualOHLT( const CONCENTRATION_DATA_ARRAY *padataConc, const int nInitIndex, COleDateTime &odtHltDateTime ): double
   // =DOC SECTION Calculate the actual HLT (hrs)
   // =DOC TEXT protected  GetActualOHLT( const CONCENTRATION_DATA_ARRAY *padataConc, const int nInitIndex, COleDateTime &odtHltDateTime ): double
   // =DOC TEXT CONCENTRATION_DATA_ARRAY *padataConc: Array of concentration values
   // =DOC TEXT int nInitIndex: Array index of the maximum concentration since last increase in concentration
   // =DOC TEXT COleDateTime &odtHltDateTime: Date/time of the final value used in calculating the HLT value
   // =DOC TEXT Returns double: Actual OHLT (hrs)
   // =DOC TEXT Calculate the actual OHLT.  Use ::GetEHLT() to convert OHLT to EHLT.
   // =DOC TEXT The algorithm is following Part II, step 2 equations as of 2001-07-25.
   //////////////////////////////////////////////////////////////////
   double GetActualOHLT(const CONCENTRATION_DATA_ARRAY* padataConc,
                        const int nInitIndex,
                        COleDateTime& odtHLTDateTime);


   // =DOC SECTION Method GetRemainingCT(const double& dCurrentConcetration, const COleDateTimeSpan& odtsCurrentTime, const COleDateTimeSpan& odtsTotalFumigationTime, const double& dActualHLT): double
   // =DOC SECTION Calculate the remaining dose.
   // =DOC TEXT protected  GetRemainingCT(const double& dCurrentConcetration, const COleDateTimeSpan& odtsCurrentTime, const COleDateTimeSpan& odtsTotalFumigationTime, const double& dActualHLT): double
   // =DOC TEXT dCurrentConcentration: The current concentration
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT odtsTotalFumigationTime: The total expected exposure time.
   // =DOC TEXT dActualOHLT: The actual OHLT at the current time, as given by GetActualOHLT().
   // =DOC TEXT Returns double: The remaining CT yet to be accumulated.
   // =DOC TEXT The algorithm is following Part II, Step 3, on 2001-11-20.
   //////////////////////////////////////////////////////////////////
   double GetRemainingCT(const double& dCurrentConcentration,
                         const COleDateTimeSpan& odtsCurrentTime,
                         const COleDateTimeSpan& odtsTotalFumigationTime,
                         const double& dActualOHLT);


   // =DOC SECTION Method GetRemainingCT2(const double& dCurrentConcetration, const COleDateTimeSpan& odtsCurrentTime, const COleDateTimeSpan& odtsTotalFumigationTime): double
   // =DOC SECTION Calculate the remaining dose using alternate method.
   // =DOC TEXT protected  GetRemainingCT2(const double& dCurrentConcetration, const COleDateTimeSpan& odtsCurrentTime, const COleDateTimeSpan& odtsTotalFumigationTime): double
   // =DOC TEXT dCurrentConcentration: The current concentration
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT odtsTotalFumigationTime: The total expected exposure time.
   // =DOC TEXT Returns double: The remaining CT yet to be accumulated.
   // =DOC TEXT The algorithm is following Part II, Step 3, on 2001-12-21.
   //////////////////////////////////////////////////////////////////
   double GetRemainingCT2(const double& dCurrentConcentration,
                          const COleDateTimeSpan& odtsCurrentTime,
                          const COleDateTimeSpan& odtsTotalFumigationTime);


   // =DOC SECTION Method GetProjectedCT(const double& dAchievedCT, const double& dRemainingCT): double
   // =DOC SECTION Get the projected CT.
   // =DOC TEXT protected GetProjectedCT(const double& dAchievedCT, const double& dRemainingCT): double
   // =DOC TEXT dAchievedCT: The currently achieved CT, as given by GetCTAchieved().
   // =DOC TEXT dRemainingCT: The CT left to accumulate, as given by GetRemainingCT().
   // =DOC TEXT Returns double: The projected final CT.
   // =doc text The algorithm is following Part II, Step 4, on 2001-07-25.
   double GetProjectedCT(const double& dAchievedCT, const double& dRemainingCT);


   // =DOC SECTION Method GetDeltaDose(const double& dTargetDose, const double& dAchievedDose): double
   // =DOC SECTION Get the Delta Dose.
   // =DOC TEXT protected GetDeltaDose(const double& dTargetDose, const double& dAchievedDose): double
   // =DOC TEXT dTargetDose: The target dosage, as given by ::GetFinalDose().
   // =DOC TEXT dAchievedDose: The dose currently achieved, as given by GetDoseAchieved().
   // =DOC TEXT Returns double: The difference between the two dosages.
   // =DOC TEXT The algorithm is following Part II, Step 7, on 2001-07-26.
   double GetDeltaDose(const double& dTargetDose, const double& dAchievedDose);


   // JNM 04/16/03 - Fum. Enh. 2003 - Added new parameter to hold pest n value.
   //		This was done so that the interpolated n value can be passed.
   // =DOC SECTION Method GetProjectedFumigationTime(const double& dActualEHLT, const double& dCurrentConcentration, const double& dDeltaDose): double
   // =DOC SECTION Determines the projected total fumigation time.
   // =DOC TEXT protected GetProjectedFumigationTime(const double& dActualEHLT, const double& dCurrentConcentration, const double& dDeltaDose): double
   // =DOC TEXT dActualEHLT: The EHLT, given by ::GetEHLT().
   // =DOC TEXT dCurrentConcentration: The current concentration.
   // =DOC TEXT dDeltaDose: The delta dose, as given by GetDeltaDose().
   // =DOC TEXT dPestNValue: The n value of the current pest
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT Returns double: The total amount of time (from exposure) needed to complete the fumigation without adding fumigant.
   // =DOC TEXT The algorithm follows Part II Step 7, as of 2001-12-20.
   double GetProjectedFumigationTime( const double& dActualEHLT,
                                      const double& dCurrentConcentration,
                                      const double& dDeltaDose,
									  const double& dPestNValue, 
                                      const COleDateTimeSpan& odtsCurrentTime );

   // JNM 04/16/03 - Fum. Enh. 2003 - Added new argument to hold n pest value
   // =DOC SECTION Method GetProjectedFumigationTime2(const double& dCurrentConcentration, const double& dDeltaDose): double
   // =DOC SECTION Determines the projected total fumigation time using an alternate method.
   // =DOC TEXT protected GetProjectedFumigationTime2(const double& dCurrentConcentration, const double& dDeltaDose): double
   // =DOC TEXT dCurrentConcentration: The current concentration.
   // =DOC TEXT dDeltaDose: The delta dose, as given by GetDeltaDose().
   // =DOC TEXT dPestNValue: the current pest's n value
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT Returns double: The total amount of time (from exposure) needed to complete the fumigation without adding fumigant.
   // =DOC TEXT The algorithm follows Part II Step 7, as of 2001-12-21.
   double GetProjectedFumigationTime2( const double& dCurrentConcentration,
                                       const double& dDeltaDose,
									   const double& dPestNValue, 
                                       const COleDateTimeSpan& odtsCurrentTime );


   // =DOC SECTION Method GetFutureConcentration( const double& dCurrentConcentration, const double& dActualHLT ): double
   // =DOC SECTION Determine the projected future concentration.
   // =DOC TEXT protected GetFutureConcentration( const double& dCurrentConcentration, const double& dActualHLT ): double
   // =DOC TEXT Determines the projected future concentration at time T+1 hours, where the current time is T.
   // =DOC TEXT dCurrentConcentration: The last maximum concentration from the current time T.
   // =DOC TEXT dActualHLT: The actual HLT
   // =DOC TEXT Returns double: The projected future concentration at time T+1 hours.
   // =DOC TEXT The algorithm follows Part II, Step 9B, as of 2001-11-20.
   double GetFutureConcentration( const double& dCurrentConcentration, const double& dActualHLT );


   // =DOC SECTION Method GetFutureDose( const double& dFutureConcentration, const double& dCurrentConcentration ): double
   // =DOC SECTION Determines the projected future dose.
   // =DOC TEXT protected GetFutureDose( const double& dFutureConcentration, const double& dCurrentConcentration ): double
   // =DOC TEXT Determines the projected future dose at time T + 1 hours, where the current time is T.
   // =DOC TEXT dFutureConcentration: The future concentration 1 hour from the current time.
   // =DOC TEXT dCurrentConcentration: The current concentration at the current time.
   // =DOC TEXT Returns double: The projected future dose at current time + 1 hour.
   // =DOC TEXT The algorithm follows Part II, Step 8C, as of 2001-11-26.
   double GetFutureDose( const double& dFutureConcentration, const double& dCurrentConcentration );


   // =DOC SECTION Method GetRemainingRequiredDose( const double& dTargetDose, const double& dAchievedDose, const double& dFutureDose ): double
   // =DOC SECTION Determines the remaining required dose.
   // =DOC TEXT protected GetRemainingRequiredDose( const double& dTargetDose, const double& dAchievedDose, const double& dFutureDose ): double
   // =DOC TEXT Determines the remaining required dose.
   // =DOC TEXT dTargetDose: The target dose for the fumigation.
   // =DOC TEXT dAchieveDose: The currently achieved dose, given by GetDoseAchieved().
   // =DOC TEXT dFutureDose: The projected future dose, given by GetFutureDose().
   // =DOC TEXT Returns double: The remaining dose to be accumulated.
   // =DOC TEXT The alogirthm follows Part II, Step 8D, on 2001-07-30.
   double GetRemainingRequiredDose( const double& dTargetDose,
                                    const double& dAchievedDose,
                                    const double& dFutureDose );


   // =DOC SECTION Method GetNewTargetConcentration( const double& dRemainingReqDose, const double& dActualEHLT, const double& dTotalFumigationTime ): double
   // =DOC SECTION Determines the new target concentration.
   // =DOC TEXT protected GetNewTargetConcentration( const double& dRemainingReqDose, const double& dActualEHLT, const double& dTotalFumigationTime ): double
   // =DOC TEXT Determines the new target concentration.
   // =DOC TEXT dRemainingReqDose: The remaining required dose given by GetRemainingRequiredDose().
   // =DOC TEXT dActualEHLT: The EHLT, given by ::GetEHLT().
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT odtsTotalFumigationTime: The total expected exposure time.
   // =DOC TEXT Returns double: The new target concentration.
   // =DOC TEXT The algorithm follows Part II, Step 9E, on 2002-01-15.
   double GetNewTargetConcentration( const double& dRemainingReqDose, 
                                     const double& dActualEHLT,
                                     const COleDateTimeSpan& odtsCurrentTime,
                                     const COleDateTimeSpan& odtsTotalFumigationTime );


   // =DOC SECTION Method GetNewTargetConcentration2( const double& dRemainingReqDose, const double& dTotalFumigationTime ): double
   // =DOC SECTION Determines the new target concentration using an alternate method.
   // =DOC TEXT protected GetNewTargetConcentration2( const double& dRemainingReqDose, const double& dTotalFumigationTime ): double
   // =DOC TEXT Determines the new target concentration using an alternate method.
   // =DOC TEXT dRemainingReqDose: The remaining required dose given by GetRemainingRequiredDose().
   // =DOC TEXT odtsCurrentTime: The current time from start of exposure.
   // =DOC TEXT odtsTotalFumigationTime: The total expected exposure time.
   // =DOC TEXT Returns double: The new target concentration.
   // =DOC TEXT The algorithm follows Part II, Step 9E, on 2002-01-15.
   double GetNewTargetConcentration2( const double& dRemainingReqDose, 
                                      const COleDateTimeSpan& odtsCurrentTime,
                                      const COleDateTimeSpan& odtsTotalFumigationTime );


   // =DOC SECTION Member m_pestCurrent: PEST_DATA
   // =DOC SECTION Structure containing the current pest data for a selected temperature
   // =DOC TEXT m_pestCurrent: PEST_DATA
   // =DOC TEXT Structure containing the current pest data for a selected temperature
   //////////////////////////////////////////////////////////////////
   PEST_DATA m_pestCurrent;


   // =DOC SECTION Member m_dataPest: DDasPestData
   // =DOC SECTION Class used to populate m_pestCurrent with the selected pest data
   // =DOC TEXT m_dataPest: DDasPestData
   // =DOC TEXT Class used to populate m_pestCurrent with the selected pest data
   //////////////////////////////////////////////////////////////////
   DDasPestData m_DasPestData;

   
   // Utility functions

   // =DOC SECTION Method GetPeakPattern( const CONCENTRATION_DATA_ARRAY *padataConc, const int nInitIndex ): double
   // =DOC SECTION Calculate the actual HLT (hrs)
   // =DOC TEXT protected  GetPeakPattern( const CONCENTRATION_DATA_ARRAY *padataConc, const int nInitIndex ): double
   // =DOC TEXT CONCENTRATION_DATA_ARRAY *padataConc: Array of concentration values
   // =DOC TEXT int nPeakIndex: Array index of the maximum concentration since last increase in concentration
   // =DOC TEXT Returns double: A numerical code indicating which of 3 possible peak patterns is occurring.
   // =DOC TEXT Calculate which of 3 possible patterns is the current occurrence:
   // =DOC TEXT    (1) The current concentration reading is increasing, so there is no HLT yet.
   // =DOC TEXT    (2) The current concentration reading is equal to the previous peak's
   // =DOC TEXT        reading, so the HLT is not measurable - effectively infinite.
   // =DOC TEXT    (3) The current concentration reading is less than a previous peak, so the
   // =DOC TEXT        HLT can be calculated following Part II, Step 2 equations.
   //////////////////////////////////////////////////////////////////
   PEAK_PATTERN GetPeakPattern(const CONCENTRATION_DATA_ARRAY* padataConc,
                               const int nPeakIndex );

   
   // =DOC SECTION Method GetLastPeakIndex( const CONCENTRATION_DATA_ARRAY *padataConc, COleDateTime &odtLastPeakDateTime ): int
   // =DOC SECTION Get the array index of the last peak concentration
   // =DOC TEXT protected  GetLastPeakIndex( const CONCENTRATION_DATA_ARRAY *padataConc, COleDateTime &odtLastPeakDateTime ): int
   // =DOC TEXT CONCENTRATION_DATA_ARRAY *padataConc: Array of concentration values
   // =DOC TEXT COleDateTime &odtLastPeakDateTime: Date/time of the peak concentration
   // =DOC TEXT Returns int: Array index of last peak concentration
   // =DOC TEXT Get the array index of the last peak concentration
   //////////////////////////////////////////////////////////////////
   static int GetLastPeakIndex( const CONCENTRATION_DATA_ARRAY *padataConc, COleDateTime &odtLastPeakDateTime );
};

#endif // !defined(AFX_MONITORPOINT_H__572A8864_59A5_4D4E_8411_6CBC31244F93__INCLUDED_)

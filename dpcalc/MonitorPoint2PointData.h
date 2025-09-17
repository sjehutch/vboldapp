// =DOC FILE Structure for monitor point 2 point information
// =DOC TEXT Declares tagMonitorPoint2PointData structure
//////////////////////////////////////////////////////////////////

// MonitorPoint2PointData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_MONITORPOINT2POINTDATA_H_)
#define _MONITORPOINT2POINTDATA_H_


#include <afxtempl.h>
#include "ConcentrationData.h"



// =DOC SECTION Struct tagMonitorPoint2PointData ( MONITOR_POINT_2_POINT_DATA, *LPMONITOR_POINT_2_POINT_DATA )
// =DOC SECTION Carrier structure containing monitoring point to point information, both input and calculated
// =DOC TEXT struct tagMonitorPoint2PointData extends nothing
// =DOC TEXT Carrier structure containing monitoring point to point information, both input and calculated
//////////////////////////////////////////////////////////////////
typedef struct tagMonitorPoint2PointData
{
   // =DOC SECTION Member padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC SECTION Array of concentration data
   // =DOC TEXT padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC TEXT Array of concentration data
   //////////////////////////////////////////////////////////////////
   CONCENTRATION_DATA_ARRAY* padataConc;


   // =DOC SECTION Member dActualOHLT: double
   // =DOC SECTION HLT for the point
   // =DOC TEXT dActualOHLT: double
   // =DOC TEXT HLT for the point
   //////////////////////////////////////////////////////////////////
   double dActualOHLT;


   // =DOC SECTION Member odtLastPeakDateTime: COleDateTime
   // =DOC SECTION Date/time for the last peak concentration
   // =DOC TEXT odtLastPeakDateTime: COleDateTime
   // =DOC TEXT Date/time for the last peak concentration
   //////////////////////////////////////////////////////////////////
   COleDateTime odtLastPeakDateTime;


   // =DOC SECTION Member odtHltDateTime: COleDateTime
   // =DOC SECTION Date/time for the concentration used to calculate last peak HLT
   // =DOC TEXT odtHltDateTime: COleDateTime
   // =DOC TEXT Date/time for the concentration used to calculate last peak HLT
   //////////////////////////////////////////////////////////////////
   COleDateTime odtHltDateTime;

   
   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} MONITOR_POINT_2_POINT_DATA, *LPMONITOR_POINT_2_POINT_DATA;

#endif // !defined(_MONITORPOINT2POINTDATA_H_)

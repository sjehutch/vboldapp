// =DOC FILE Structure for pest type information
// =DOC TEXT Declares tagPestData structure, pest enumerations
//////////////////////////////////////////////////////////////////

// PestData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_PESTDATA_H_)
#define _PESTDATA_H_


#include <afxtempl.h>



// =DOC SECTION Struct tagPestData ( PEST_DATA, *LPPEST_DATA )
// =DOC SECTION Carrier structure containing pest information
// =DOC TEXT struct tagPestData extends nothing
// =DOC TEXT Carrier structure containing pest information
//////////////////////////////////////////////////////////////////
typedef struct tagPestData
{
   // =DOC SECTION Member dTemp: double
   // =DOC SECTION Temperature for pest data (C)
   // =DOC TEXT dTemp: double
   // =DOC TEXT Temperature for pest data (C)
   //////////////////////////////////////////////////////////////////
   double dTemp;

   // =DOC SECTION Member dNValue: double
   // =DOC SECTION N value
   // =DOC TEXT dNValue: double
   // =DOC TEXT N value
   //////////////////////////////////////////////////////////////////
   double dNValue;

   // =DOC SECTION Member dIntercept: double
   // =DOC SECTION Intercept value
   // =DOC TEXT dIntercept: double
   // =DOC TEXT Intercept value
   //////////////////////////////////////////////////////////////////
   double dIntercept;

   // =DOC SECTION Member dSlope: double
   // =DOC SECTION Slope value
   // =DOC TEXT dSlope: double
   // =DOC TEXT Slope value
   //////////////////////////////////////////////////////////////////
   double dSlope;

   // =DOC SECTION Member dProbit: double
   // =DOC SECTION Probit value
   // =DOC TEXT dProbit: double
   // =DOC TEXT Probit value
   //////////////////////////////////////////////////////////////////
   double dProbit;

   // =DOC SECTION Member dSafetyFactor: double
   // =DOC SECTION Safety factor value
   // =DOC TEXT dSafetyFactor: double
   // =DOC TEXT Safety factor value
   //////////////////////////////////////////////////////////////////
   double dSafetyFactor;

   // =DOC SECTION struct constructor
   // =DOC SECTION Initializes the structure
   // =DOC TEXT struct constructor extends nothing
   // =DOC TEXT Initializes the structure.
   //////////////////////////////////////////////////////////////////
   tagPestData()
   {
      dTemp = 0.0;
      dNValue = 0.0;
      dIntercept = 0.0;
      dSlope = 0.0;
      dProbit = 0.0;
      dSafetyFactor = 0.0;
   };

} PEST_DATA, *LPPEST_DATA;


// =DOC SECTION CArray PEST_DATA_ARRAY
// =DOC SECTION Array of PEST_DATA structures
// =DOC TEXT CArray PEST_DATA_ARRAY extends CArray
// =DOC TEXT Array of PEST_DATA structures
//////////////////////////////////////////////////////////////////
typedef CArray < PEST_DATA, PEST_DATA& > PEST_DATA_ARRAY;

#endif // !defined(_PESTDATA_H_)
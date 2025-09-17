// =DOC FILE Structure for commodity type information
// =DOC TEXT Declares tagCommodityData structure, commodity and fumigation type enumerations
//////////////////////////////////////////////////////////////////

// CommodityData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_COMMODITYDATA_H_)
#define _COMMODITYDATA_H_


#include <afxtempl.h>



// =DOC SECTION Struct tagPestData ( PEST_DATA, *LPPEST_DATA )
// =DOC SECTION Carrier structure containing pest information
// =DOC TEXT struct tagPestData extends nothing
// =DOC TEXT Carrier structure containing pest information
//////////////////////////////////////////////////////////////////
typedef struct tagCommodityData
{
   // =DOC SECTION Member dMaxConcentrationTime: double
   // =DOC SECTION Maximum permitted CT
   // =DOC TEXT dMaxConcentrationTime: double
   // =DOC TEXT Maximum permitted CT
   //////////////////////////////////////////////////////////////////
   double dMaxConcentrationTime;

   //PAT 1/13/06 - Added
   // =DOC SECTION Member dPorosity: double
   // =DOC SECTION Commodity's porosity
   // =DOC TEXT dPorosity: double
   // =DOC TEXT Commodity's porosity
   //////////////////////////////////////////////////////////////////
   double dPorosity; 

   //PAT 1/13/06 - Added
   // =DOC SECTION Member dBConstant: double
   // =DOC SECTION Commodity's B Constant
   // =DOC TEXT dBConstant: double
   // =DOC TEXT Commodity's B Constant
   //////////////////////////////////////////////////////////////////
   double dBConstant; 

   //PAT 1/13/06 - Added
   // =DOC SECTION Member dCConstant: double
   // =DOC SECTION Commodity's C Constant
   // =DOC TEXT dCConstant: double
   // =DOC TEXT Commodity's C Constant
   //////////////////////////////////////////////////////////////////
   double dCConstant; 

   //PAT 1/13/06 - Added
   // =DOC SECTION Member dDConstant: double
   // =DOC SECTION Commodity's D Constant
   // =DOC TEXT dDConstant: double
   // =DOC TEXT Commodity's D Constant
   //////////////////////////////////////////////////////////////////
   double dDConstant; 

   //PAT 1/13/06 - Added
   // =DOC SECTION Member dDCZero: double
   // =DOC SECTION Commodity's DC_Zero Constant
   // =DOC TEXT dDCZero: double
   // =DOC TEXT Commodity's DC_Zero Constant
   //////////////////////////////////////////////////////////////////
   double dDCZero; 

   // =DOC SECTION struct constructor
   // =DOC SECTION Initializes the structure
   // =DOC TEXT struct constructor extends nothing
   // =DOC TEXT Initializes the structure.
   //////////////////////////////////////////////////////////////////
   tagCommodityData()
   {
      dMaxConcentrationTime = 0.0;
	  dPorosity = 0.0;
	  dBConstant = 0.0;
	  dCConstant = 0.0;
	  dDConstant = 0.0;
	  dDCZero = 0.0;
   };
} COMMODITY_DATA, *LPCOMMODITY_DATA;


// =DOC SECTION CArray COMMODITY_DATA_ARRAY
// =DOC SECTION Array of COMMODITY_DATA structures
// =DOC TEXT CArray COMMODITY_DATA_ARRAY extends CArray
// =DOC TEXT Array of COMMODITY_DATA structures
//////////////////////////////////////////////////////////////////
typedef CArray < COMMODITY_DATA, COMMODITY_DATA& > COMMODITY_DATA_ARRAY;


#endif // !defined(_COMMODITYDATA_H_)
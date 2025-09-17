// =DOC FILE Interface for the DDoseCalc class
// =DOC TEXT Declares DDoseCalc class to determine the amount of profume required for a fumigation
//////////////////////////////////////////////////////////////////

// DoseCalc.h: interface for the DDoseCalc class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DOSECALC_H__B421F503_5A56_4F59_A5A6_00DE5C433D77__INCLUDED_)
#define AFX_DOSECALC_H__B421F503_5A56_4F59_A5A6_00DE5C433D77__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "CommodityData.h"
#include "DasCommodityData.h"
#include "DasPestData.h"
#include "DoseData.h"
#include "PestData.h"


// =DOC SECTION Class DDoseCalc
// =DOC SECTION Determine how much ProFume is required for a fumigation
// =DOC TEXT class DDoseCalc extends nothing
// =DOC TEXT Determine how much ProFume is required for a fumigation
//////////////////////////////////////////////////////////////////
class DDoseCalc  
{
public:

   // =DOC SECTION Method GetProfumeRequired( LPDOSE_DATA lpdataDose ): BOOL
   // =DOC SECTION Get the amount of ProFume required for a fumigation
   // =DOC TEXT public  GetProfumeRequired( LPDOSE_DATA lpdataDose ): BOOL
   // =DOC TEXT LPDOSE_DATA lpdataDose: A carrier structure for the fumigation
   // =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Get the amount of ProFume required and other requied data for a fumigation.  
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
   BOOL GetProfumeRequired( LPDOSE_DATA lpdataDose );

   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
   DDoseCalc();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DDoseCalc();


protected:
   // =DOC SECTION Method GetConcentrationTime( const double &dHltOriginal, const double &dTimeExposure, const double &dInitialConcentration ): double
   // =DOC SECTION Get the CT value.
   // =DOC TEXT protected  GetConcentrationTime( const double &dHltOriginal, const double &dTimeExposure, const double &dInitialConcentration ): double
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEST double &dTimeExposure: User entered time of exposure
   // =DOC TEST double &dInitialConcentration: The initial maximum concentration
   // =DOC TEXT Returns double: Returns the CT value
   // =DOC TEXT Get the CT value
   //////////////////////////////////////////////////////////////////
   double GetConcentrationTime( const double &dHltOriginal, 
                                const double &dTimeExposure, 
                                const double &dInitialConcentration );
   
   // =DOC SECTION Method GetInitConcentration( const double &dHltOriginal, const double &dTimeExposure, const double &dFinalDose ): double
   // =DOC SECTION Get the initial concentration value based on the final dose.
   // =DOC TEXT protected  GetInitConcentration( const double &dHltOriginal, const double &dTimeExposure, const double &dFinalDose ): double
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEST double &dTimeExposure: User entered time of exposure
   // =DOC TEST double &dFinalDose: The final dose value
   // =DOC TEST bool dUserDefined: Flag that says if user has defined a CT or not.12/23/04-YYT
   // =DOC TEST bool dVolume: Volume that is entered by the user.01/27/06-PAT
   // =DOC TEST bool dTemp: Temperature that is entered by the user.01/27/06-PAT
   // =DOC TEST bool dBConstant: B constant of a commodity.01/27/06-PAT
   // =DOC TEST bool dCConstant: C constant of a commodity.01/27/06-PAT
   // =DOC TEST bool dCDonstant: C constant of a commodity.01/27/06-PAT
   // =DOC TEST bool dPorosity: Porosity constant of a commodity.01/27/06-PAT
   // =DOC TEST bool dDCZero: DCZero constant of a commodity.01/27/06-PAT
   // =DOC TEST bool dLoadFactor: Load Factor that is entered by the user.01/27/06-PAT
   // =DOC TEST bool IsComplete: Will indicate if the criteria is complete for the comp of Adj HLT.01/27/06-PAT
   // =DOC TEST bool nCounter: Will state if we are on what record.01/27/06-PAT
   // =DOC TEXT Returns double: Returns the initial concentration value
   // =DOC TEXT Get the initial concentration value based on the final dose.
   //////////////////////////////////////////////////////////////////
   double GetInitConcentration( const double &dHltOriginal, 
                                const double &dTimeExposure, 
                                const double &dFinalDose,
								bool dUserDefined,
							    const double &dVolume,
								const double &dTemp,
								const double &dBConstant,
								const double &dCConstant,
								const double &dCDonstant,
								const double &dPorosity,
								const double &dDCZero,
								const double &dLoadFactor,
								const int IsComplete,
								const int nCounter);
	

   // =DOC SECTION Method GetEstimatedHlt( const double &dHltOriginal ): double
   // =DOC SECTION Get the estimated HLT value.
   // =DOC TEXT protected  GetEstimatedHlt( const double &dHltOriginal ): double
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEXT Returns double: Returns the estimated HLT value
   // =DOC TEXT Get the estimated HLT value
   //////////////////////////////////////////////////////////////////
   double GetEstimatedHlt( const double &dHltOriginal );


   // =DOC SECTION Method GetMaxInitConcentration( const double &dHltOriginal, const double &dTimeExposure ): double
   // =DOC SECTION Get the maximum initial concentration value.
   // =DOC TEXT protected  GetMaxInitConcentration( const double &dHltOriginal, const double &dTimeExposure ): double
   // =DOC TEXT double &dMaxConcentrationTime: Maximum permitted concentration time value
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEST double &dTimeExposure: User entered time of exposure
   // =DOC TEXT Returns double: Returns the maximum initial concentration value
   // =DOC TEXT Get the maximum initial concentration value
   //////////////////////////////////////////////////////////////////
   double GetMaxInitConcentration( const double &dMaxConcentrationTime,
                                   const double &dHltOriginal, 
                                   const double &dTimeExposure );

   // =DOC SECTION Method GetMaxDose( const double &dHltOriginal, const double &dTimeExposure, const double &dInitialConcentration ): double
   // =DOC SECTION Get the maximum dose value based on the max initial concentration
   // =DOC TEXT protected  GetMaxDose( const double &dHltOriginal, const double &dTimeExposure, const double &dInitialConcentration ): double
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEST double &dTimeExposure: User entered time of exposure
   // =DOC TEST double &dInitialConcentration: The initial maximum concentration
   // =DOC TEXT Returns double: Returns the maximum dose value
   // =DOC TEXT Get the maximum dose value based on the max initial concentration
   //////////////////////////////////////////////////////////////////
   double GetMaxDose( const double &dHltOriginal, 
                      const double &dTimeExposure, 
                      const double &dInitialConcentration,
					  const int IsComplete,
					  bool dUserDefined);

   // =DOC SECTION Method GetCommodityDisplacement( const double &dVolume, const double &dLoadFactor, const double &dPorosity ): double
   // =DOC SECTION Get the commodity displacement
   // =DOC TEXT protected  GetCommodityDisplacement( const double &dVolume, const double &dLoadFactor, const double &dPorosity ): double
   // =DOC TEXT double &dVolume: User entered Volume value
   // =DOC TEST double &dLoadFactor: User entered Load Factor of exposure
   // =DOC TEST double &dPorosity: Porosity of the chosen commodity
   // =DOC TEXT Returns double: Returns the commodity displacement
   // =DOC TEXT Get the commodity displacement
   //////////////////////////////////////////////////////////////////
   double GetCommodityDisplacement ( const double &dVolume, 
									 const double &dLoadFactor, 
									 const double &dPorosity);


   // =DOC SECTION Method GetAlphaBasedonHLT0( const double &dHltOriginal ): double
   // =DOC SECTION Get the Alpha Based on HLT0
   // =DOC TEXT protected  GetAlphaBasedonHLT0( const double &dHltOriginal ): double
   // =DOC TEXT double &dHltOriginal: User entered HLT value
   // =DOC TEXT Returns double: Returns the Alpha Based on HLT0
   // =DOC TEXT Get the Alpha Based on HLT0
   //////////////////////////////////////////////////////////////////
   double GetAlphaBasedonHLT0 ( const double &dHltOriginal );

   // =DOC SECTION Method GetBetaSorption( const double &dTemp, const double &dBConstant, const double &dCConstant, const double &dDCZero ): double
   // =DOC SECTION Get the Beta Sorption
   // =DOC TEXT protected  GetBetaSorption( const double &dTemp, const double &dBConstant, const double &dCConstant, const double &dDCZero ): double
   // =DOC TEXT double &dTemp: User entered Temp value
   // =DOC TEXT double &dBConstant: Constant for the chosen commodity
   // =DOC TEST double &dCConstant: Constant for the chosen commodity
   // =DOC TEST double &dDCZero: Constant for the chosen commodity
   // =DOC TEXT Returns double: Returns the Beta Sorption
   // =DOC TEXT Get the Beta Sorption
   //////////////////////////////////////////////////////////////////
   double GetBetaSorption ( const double &dTemp,
						   const double &dBConstant,
						   const double &dCConstant,
						   const double &dDCZero );

   // =DOC SECTION Method GetAlphaSorption( const double &dLoadFactor ): double
   // =DOC SECTION Get the Alpha Sorption
   // =DOC TEXT protected  GetAlphaSorption( const double &dLoadFactor ): double
   // =DOC TEXT double &dHltOriginal: User entered Load Factor value
   // =DOC TEXT Returns double: Returns the Alpha Sorption
   // =DOC TEXT Get the Alpha Sorption
   //////////////////////////////////////////////////////////////////
   double GetAlphaSorption ( const double &dLoadFactor,
							 const double dBetaSorpt);

   // =DOC SECTION Method GetAlphaSorption( const double &dLoadFactor ): double
   // =DOC SECTION Get the Alpha Sorption
   // =DOC TEXT protected  GetAlphaSorption( const double &dLoadFactor ): double
   // =DOC TEXT double &dHltOriginal: User entered Load Factor value
   // =DOC TEXT Returns double: Returns the Alpha Sorption
   // =DOC TEXT Get the Alpha Sorption
   //////////////////////////////////////////////////////////////////
   double GetBetaSorptionPrime ( const double &dTemp,
						         const double &dBConstant,
						         const double &dCConstant,
						         const double &dDConstant,
						         const double dCnTTarget);



private:

	// =DOC SECTION Member m_dCurrentTemp: double
	// =DOC SECTION Contains the current temperature of the fumigation
	// =DOC TEXT m_dCurrentTemp: double
	// =DOC TEXT Contains the current temperature of the fumigation
	//////////////////////////////////////////////////////////////////
	double m_dCurrentTemp;


   // =DOC SECTION Member m_pestCurrent: PEST_DATA
   // =DOC SECTION Structure containing the current pest data for a selected temperature
   // =DOC TEXT m_pestCurrent: PEST_DATA
   // =DOC TEXT Structure containing the current pest data for a selected temperature
   //////////////////////////////////////////////////////////////////
   PEST_DATA m_pestCurrent;


   // =DOC SECTION Member m_dasPestData: DDasPestData
   // =DOC SECTION Class used to populate m_pestCurrent with the selected pest data
   // =DOC TEXT m_dasPestData: DDasPestData
   // =DOC TEXT Class used to populate m_pestCurrent with the selected pest data
   //////////////////////////////////////////////////////////////////
   DDasPestData m_dasPestData;


   // =DOC SECTION Member m_dataCommodity: COMMODITY_DATA
   // =DOC SECTION Structure containing the current commodity data for a selected treatment type
   // =DOC TEXT m_dataCommodity: COMMODITY_DATA
   // =DOC TEXT Structure containing the current commodity data for a selected treatment type
   //////////////////////////////////////////////////////////////////
   COMMODITY_DATA m_dataCommodity;


   // =DOC SECTION Member m_dasCommodityData: DDasCommodityData
   // =DOC SECTION Class used to populate m_dataCommodity with the selected commodity data include allowable CT
   // =DOC TEXT m_dasCommodityData: DDasCommodityData
   // =DOC TEXT Class used to populate m_dataCommodity with the selected commodity data include allowable CT
   //////////////////////////////////////////////////////////////////
   DDasCommodityData m_dasCommodityData;


   //PAT 1/25/06 - Added
   // =DOC SECTION Member dAdjHLT: double
   // =DOC SECTION This will hold the Adj HLT due to Sorption
   // =DOC TEXT dAdjHLT: double
   // =DOC TEXT This will hold the Adj HLT due to Sorption
   //////////////////////////////////////////////////////////////////
   double dAdjHLT;

 
};

#endif // !defined(AFX_DOSECALC_H__B421F503_5A56_4F59_A5A6_00DE5C433D77__INCLUDED_)

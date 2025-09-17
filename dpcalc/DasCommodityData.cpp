// CommodityData.cpp: implementation of the DCommodityData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "DasCommodityData.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDasCommodityData::DDasCommodityData()
{

}

DDasCommodityData::~DDasCommodityData()
{

}


void DDasCommodityData::GetCommodityNames( CStringArray &astrCommodityNames )
{
   // Add all of the commodities this calculator supports
   // Ensure the commodity table is empty
   astrCommodityNames.RemoveAll();
   
   // Ensure the order that the commodities appear in the array are consistant
   // with the enum COMMODITY.  See CommodityData.h for the order.

   CString strCommodity;
   strCommodity.LoadString( IDS_NONE );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_ALMONDS );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_FIGS );
   astrCommodityNames.Add( strCommodity );

   // Pat 1/16/06 - Removed FILBERTS_HAZELNUTS changed to APRICOTS_DRIED
   strCommodity.LoadString( IDS_APRICOTS_DRIED );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_PISTACHIOS );
   astrCommodityNames.Add( strCommodity );

   // Pat 1/16/06 - Removed PRUNES changed to CACAO
   strCommodity.LoadString( IDS_CACAO );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_RAISINS );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_WALNUTS );
   astrCommodityNames.Add( strCommodity );

   // Pat 1/16/06 - Removed WHOLE_GRAIN changed to COFFEE
   strCommodity.LoadString( IDS_COFFEE );
   astrCommodityNames.Add( strCommodity );

   //JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new commodity called "Other".
   // Pat 1/16/06 - Removed OTHER changed to CORN
   strCommodity.LoadString( IDS_CORN );
   astrCommodityNames.Add( strCommodity );


   // Pat 1/16/06 - Added the following COMMODITIES
   strCommodity.LoadString( IDS_COWPEA );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_DATES );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_OTHER_DRIED_FRUIT );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_OTHER_GRAIN );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_OTHER_LEGUME );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_OTHER_NUTS );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_PECANS );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_PLUMS_DRIED );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_POPCORN );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_RICE_POLISHED );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_RICE_ROUGH );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_SOYBEAN );
   astrCommodityNames.Add( strCommodity );

   strCommodity.LoadString( IDS_WHEAT );
   astrCommodityNames.Add( strCommodity );
   //PAT 1/16/06 -  ENd of NEW Commodities
}

void DDasCommodityData::PopulateCommodityData( const COMMODITY &eCommodity, 
                                               const FUMIGATION_TYPE &eFumigationType, 
                                               const TREATMENT_TYPE &eTreatmentType, 
                                               CString &strMsg )
{
   // Clear out the existing data
   m_dataCommodity.dMaxConcentrationTime = 0.0;


   //PAT 1/13/05 - Removed this condition, even though the user has chosen SPACE Fumigation
   //it will still get the values for each commodities

   // If this is a space fumigation, get values for none
   // otherwise get the values for the appropriate commodity
   //if( eFumigationType == SPACE_FUMIGATION )
   //{
	  //JNM 04/30/03 - PT Fix - Since PopulateNone has been changed, the original
	  //	functionality here should be retained which is to set the max to 1500.
      //PopulateNone( eTreatmentType, strMsg );

   //CG 01/28/2004 Launch Changes 2004 Space/Vacuum max CT to 200 from the 1500CT 
   //CG 02/06/2004 Added condition to check if the Treatment Type is
   //              Vacuum or NAP
   //	if( eTreatmentType == VACUUM )
   //   {
   //		m_dataCommodity.dMaxConcentrationTime = 200;
   //		}
   //		else
   //		{
   //			m_dataCommodity.dMaxConcentrationTime = 1500;
   //		}
   //   }
   //   else
   //   {
   // Populate the commodity array with the selected commodity's data
      switch( eCommodity )
      {
      case COMMODITY_NONE:
         PopulateNone( eTreatmentType, strMsg );
         break;

      case COMMODITY_WALNUTS:
         PopulateWalnuts( eTreatmentType, strMsg );
         break;

      case COMMODITY_ALMONDS:
         PopulateAlmonds( eTreatmentType, strMsg );
         break;

      case COMMODITY_FIGS:
         PopulateFigs( eTreatmentType, strMsg );
         break;

	  //Pat 1/17/2006 - Removed FILBERTS_HAZELNUTS
      //case COMMODITY_FILBERTS_HAZELNUTS:
      //   PopulateFilbertsHazelnuts( eTreatmentType, strMsg );
      //   break;

      case COMMODITY_PISTACHIOS:
         PopulatePistachios( eTreatmentType, strMsg );
         break;

	  //Pat 1/17/2006 - Removed PRUNES
      //case COMMODITY_PRUNES:
      //   PopulatePrunes( eTreatmentType, strMsg );
      //   break;

      case COMMODITY_RAISINS:
         PopulateRaisins( eTreatmentType, strMsg );
         break;

	  //Pat 1/17/2006 - Removed WHOLE_GRAIN
      //case COMMODITY_WHOLE_GRAIN:
      //   PopulateWholeGrain( eTreatmentType, strMsg );
      //   break;

	  //Pat 1/17/2006 - Removed OTHER
	  //JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new case for Other commodity.
	  //case COMMODITY_OTHER:
      //   PopulateOther( eTreatmentType, strMsg );
      //   break;

	  // Pat 1/16/06 - Added the following COMMODITIES
	  case COMMODITY_APRICOTS_DRIED:
		 PopulateApricotsDried( eTreatmentType, strMsg );
         break;

	  case COMMODITY_CACAO:
		 PopulateCacao( eTreatmentType, strMsg );
         break;

	  case COMMODITY_COFFEE:
		 PopulateCoffee( eTreatmentType, strMsg );
         break;

	  case COMMODITY_CORN:
		 PopulateCorn( eTreatmentType, strMsg );
         break;

	  case COMMODITY_COWPEA:
		 PopulateCowpea( eTreatmentType, strMsg );
         break;

	  case COMMODITY_DATES:
		 PopulateDates( eTreatmentType, strMsg );
         break;

	  case COMMODITY_OTHER_DRIED_FRUIT:
		 PopulateOtherDriedFruit( eTreatmentType, strMsg );
         break;

	  case COMMODITY_OTHER_GRAIN:
		 PopulateOtherGrain( eTreatmentType, strMsg );
         break;

	  case COMMODITY_OTHER_LEGUME:
		 PopulateOtherLegume( eTreatmentType, strMsg );
         break;

	  case COMMODITY_OTHER_NUTS:
		 PopulateOtherNuts( eTreatmentType, strMsg );
         break;

	  case COMMODITY_PECANS:
		 PopulatePecans( eTreatmentType, strMsg );
         break;

	  case COMMODITY_PLUMS_DRIED:
		 PopulatePlumsDried( eTreatmentType, strMsg );
         break;

	  case COMMODITY_POPCORN:
		 PopulatePopcorn( eTreatmentType, strMsg );
         break;

	  case COMMODITY_RICE_POLISHED:
		 PopulateRicePolished( eTreatmentType, strMsg );
         break;

	  case COMMODITY_RICE_ROUGH:
		 PopulateRiceRough( eTreatmentType, strMsg );
         break;

	  case COMMODITY_SOYBEAN:
		 PopulateSoybean( eTreatmentType, strMsg );
         break;

	  case COMMODITY_WHEAT:
		 PopulateWheat( eTreatmentType, strMsg );
         break;
      //Pat 1/16/06 -  ENd of NEW Commodities

      default:
         // We should not be here
         ASSERT( FALSE );
     }
//   }
}


COMMODITY_DATA DDasCommodityData::GetCommodityData()
{
   return m_dataCommodity;
}


void DDasCommodityData::PopulateNone( const TREATMENT_TYPE &eTreatmentType, CString &strMsg )
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }
	
   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0;
   m_dataCommodity.dCConstant = 0;
   m_dataCommodity.dDConstant = 0;
   m_dataCommodity.dPorosity  = 1;
   m_dataCommodity.dDCZero = 0;

}

void DDasCommodityData::PopulateWalnuts(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; //TBD
   m_dataCommodity.dCConstant = 0; //TBD
   m_dataCommodity.dDConstant = 0; //TBD
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0; //TBD
}

void DDasCommodityData::PopulateAlmonds(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
       m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; //TBD
   m_dataCommodity.dCConstant = 0; //TBD
   m_dataCommodity.dDConstant = 0; //TBD
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0; //TBD
}

void DDasCommodityData::PopulateFigs(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00212703118367347; 
   m_dataCommodity.dCConstant = -0.000181660408163265; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}
//Pat 1/17/2006 - Removed PopulateFilbertsHazelnuts function
/*void DDasCommodityData::PopulateFilbertsHazelnuts(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
     m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }
}*/

void DDasCommodityData::PopulatePistachios(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; //TBD
   m_dataCommodity.dCConstant = 0; //TBD
   m_dataCommodity.dDConstant = 0; //TBD
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0; //TBD
}

//Pat 1/17/2006 - Removed PopulatePrunes function
/*void DDasCommodityData::PopulatePrunes(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }
}*/

void DDasCommodityData::PopulateRaisins(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; 
   m_dataCommodity.dCConstant = 0; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}
//Pat 1/17/2006 - Removed PopulateWholeGrain function
/*void DDasCommodityData::PopulateWholeGrain(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }
}*/

//Pat 1/17/2006 - Removed PopulateWholeGrain function
//JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new method for new commodity.
/*void DDasCommodityData::PopulateOther(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   //JNM 04/14/03 - Fum. Enh. 2003 - Set max CT at 200 for vacuum and 1500 for NAP
   //ASSERT( &eTreatmentType );
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }
}*/

//Pat 1/17/06 - Added PopulateApricotsDried function
void DDasCommodityData::PopulateApricotsDried(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.000840244416326531; 
   m_dataCommodity.dCConstant = -0.0000717615918367347; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 

}

//Pat 1/17/06 - Added PopulateCacao function
void DDasCommodityData::PopulateCacao(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.005026724; 
   m_dataCommodity.dCConstant = -0.0009045; 
   m_dataCommodity.dDConstant = 0.000112794394904459; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0.00328; 
}

//Pat 1/17/06 - Added PopulateCoffee function
void DDasCommodityData::PopulateCoffee(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00239148423265306; 
   m_dataCommodity.dCConstant = -0.000204246183673469; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}

//Pat 1/17/06 - Added PopulateCorn function
void DDasCommodityData::PopulateCorn(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = -0.017359; 
   m_dataCommodity.dCConstant = -0.001528; 
   m_dataCommodity.dDConstant = 0.00059447983014862; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0.018; 
}

//Pat 1/17/06 - Added PopulateCowpea function
void DDasCommodityData::PopulateCowpea(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00480927403571429; 
   m_dataCommodity.dCConstant = -0.000773236607142857; 
   m_dataCommodity.dDConstant = 0.0000860450272975435; 
   m_dataCommodity.dPorosity  = 0.33;
   m_dataCommodity.dDCZero = 0.00250214285714286; 
}

//Pat 1/17/06 - Added PopulateDates function
void DDasCommodityData::PopulateDates(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00212703118367347; 
   m_dataCommodity.dCConstant = -0.000181660408163265; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}

//Pat 1/17/06 - Added PopulateOtherDriedFruit function
void DDasCommodityData::PopulateOtherDriedFruit(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00212703118367347; 
   m_dataCommodity.dCConstant = -0.000181660408163265; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}

//Pat 1/17/06 - Added PopulateOtherGrain function
void DDasCommodityData::PopulateOtherGrain(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.005228; 
   m_dataCommodity.dCConstant = -0.001026; 
   m_dataCommodity.dDConstant = 0.000137554140127389; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0.004; 
}

//Pat 1/17/06 - Added PopulateOtherLegume function
void DDasCommodityData::PopulateOtherLegume(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00480927403571429; 
   m_dataCommodity.dCConstant = -0.000773236607142857; 
   m_dataCommodity.dDConstant = 0.0000860450272975435; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0.00250214285714286; 
}


//Pat 1/17/06 - Added PopulateOtherNuts function
void DDasCommodityData::PopulateOtherNuts(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; //TBD
   m_dataCommodity.dCConstant = 0; //TBD
   m_dataCommodity.dDConstant = 0; //TBD
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0; //TBD
}

//Pat 1/17/06 - Added PopulatePecans function
void DDasCommodityData::PopulatePecans(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0; //TBD
   m_dataCommodity.dCConstant = 0; //TBD
   m_dataCommodity.dDConstant = 0; //TBD
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0; //TBD
}

//Pat 1/17/06 - Added PopulatePlumsDried function
void DDasCommodityData::PopulatePlumsDried(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00212703118367347; 
   m_dataCommodity.dCConstant = -0.000181660408163265; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}


//Pat 1/17/06 - Added PopulatePopcorn function
void DDasCommodityData::PopulatePopcorn(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00493187667857143; 
   m_dataCommodity.dCConstant = -0.000847245535714286; 
   m_dataCommodity.dDConstant = 0.000101126856232939; 
   m_dataCommodity.dPorosity  = 0.35;
   m_dataCommodity.dDCZero = 0.00294071428571429; 
}

//Pat 1/17/06 - Added PopulateRicePolished function
void DDasCommodityData::PopulateRicePolished(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00327886553877551; 
   m_dataCommodity.dCConstant = -0.000280033530612245; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}


//Pat 1/17/06 - Added PopulateRiceRough function
void DDasCommodityData::PopulateRiceRough(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.005228; 
   m_dataCommodity.dCConstant = -0.001026; 
   m_dataCommodity.dDConstant = 0.000137554140127389; 
   m_dataCommodity.dPorosity  = 0.45;
   m_dataCommodity.dDCZero = 0.004; 
}

//Pat 1/17/06 - Added PopulateSoybean function
void DDasCommodityData::PopulateSoybean(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.00504169989285714; 
   m_dataCommodity.dCConstant = -0.000913540178571429; 
   m_dataCommodity.dDConstant = 0.000114636637852594; 
   m_dataCommodity.dPorosity  = 0.33;
   m_dataCommodity.dDCZero = 0.00333357142857143; 
}

//Pat 1/17/06 - Added PopulateWheat function
void DDasCommodityData::PopulateWheat(const TREATMENT_TYPE &eTreatmentType, CString &strMsg)
{
   ASSERT( strMsg );

   if( eTreatmentType == VACUUM )
   {
      m_dataCommodity.dMaxConcentrationTime = 200;
   }
   else
   {
      m_dataCommodity.dMaxConcentrationTime = 1500;
   }

   //Pat 1/17/06 - Added parameters for the computation of Adjustments to 
   //              Expected HLT due to Sorption
   m_dataCommodity.dBConstant = 0.0041098; 
   m_dataCommodity.dCConstant = -0.000351; 
   m_dataCommodity.dDConstant = 0; 
   m_dataCommodity.dPorosity  = 0.4;
   m_dataCommodity.dDCZero = 0; 
}

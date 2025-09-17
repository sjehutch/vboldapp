// MaxShootingRate.cpp: implementation of the DMaxShootingRate class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "GlobalEnums.h"

#include "MaxShootingRate.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DMaxShootingRate::DMaxShootingRate()
{
   // Ensure it is empty to start with
   m_adataRelHumidity.RemoveAll();
}

DMaxShootingRate::~DMaxShootingRate()
{
   // Clean up the array
   m_adataRelHumidity.RemoveAll();
}


BOOL DMaxShootingRate::GetMaxShootingRate( LPMAXSHOOTINGRATE_DATA lpdataMaxRate )
{
   BOOL bReturn = TRUE;

   PopulateRelativeHumidityIndex();

   double dRelativeHumidityIndex = GetRelativeHumidityIndex( lpdataMaxRate->nRelativeHumidity );
   
   // Get the maximum shooting rate
   double dMaxRate = GetMaxShootingRate( lpdataMaxRate->dFanCapacity, dRelativeHumidityIndex );

   if( dMaxRate > MAX_SHOOTING_RATE )
   {
      dMaxRate = MAX_SHOOTING_RATE;
   }

   if( dMaxRate < MIN_SHOOTING_RATE )
   {
      dMaxRate = MIN_SHOOTING_RATE;
      lpdataMaxRate->strError.LoadString( IDS_WARNING_MAX_SHOOTING_RATE_LOW );
   }

   lpdataMaxRate->dMaxShootingRate = dMaxRate;

   return bReturn;
}


double DMaxShootingRate::GetMaxShootingRate( const double &dFanCapacity, 
                                             const double &dRelativeHumidityIndex )
{
   // Maximum Shooting Rate calculation from Dave Cale
   //
   //       MSR = (FC / 141.983) * [(-0.23 * RH + 21.94) - 1]
   //
   // MSR = Maximum Shooting Rate, kg/min
   //  FC = Fan Capacity per 1000 cubic ft, m^3 /min
   //  RH = Relative Humidity, %

   // Simplified by Brian Schneider on 10-25-02 to the following equation
   // English Version
   //                      Fan Capacity (cuft/min)
   // Max Shooting Rate = ------------------------- * Relative Humidity Index
   //                              1000 cuf

   // Metric Version
   //                       Fan Capacity (cum/min)
   // Max Shooting Rate = ------------------------- * 0.45359237 kgs/lbs * Relative Humidity Index
   //                           28.316847 cum

   double dReturn = ( dFanCapacity / 28.316847 ) * 0.45359237 * dRelativeHumidityIndex;
   
   return dReturn;
}


double DMaxShootingRate::GetRelativeHumidityIndex(const int &nRelativeHumidity)
{
   double dReturn = 0.0;

   // Find the relative humidity that is >= to the users value
   int nTotal = m_adataRelHumidity.GetSize();
   
   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      REL_HUMIDITY_DATA& dataRelHumidityUpper = m_adataRelHumidity.ElementAt( nIndex );

      if( dataRelHumidityUpper.nRelativeHumidity >= nRelativeHumidity )
      {
         // Check if it is the exact value or if we are at the lower bounds
         if( dataRelHumidityUpper.nRelativeHumidity == nRelativeHumidity || nIndex == 0 )
         {
            dReturn = dataRelHumidityUpper.dIndex;
            break;
         }
         else
         {
            // We need to get the previous value and interpolate between the upper and lower values
            REL_HUMIDITY_DATA& dataRelHumidityLower = m_adataRelHumidity.ElementAt( nIndex - 1 );
         
            double dIntRelHumidity = static_cast< double >( nRelativeHumidity - dataRelHumidityLower.nRelativeHumidity ) / 
                                     static_cast< double >( dataRelHumidityUpper.nRelativeHumidity - dataRelHumidityLower.nRelativeHumidity );

            dReturn = ( dataRelHumidityUpper.dIndex - dataRelHumidityLower.dIndex ) * dIntRelHumidity + dataRelHumidityLower.dIndex;

            break;
         }
      }
   }
   
   return dReturn;
}


void DMaxShootingRate::PopulateRelativeHumidityIndex()
{
   REL_HUMIDITY_DATA dataRelHumidity;

   // Populate the relative humidity data array with the data
   // Currently all index values are 1.0.  This decision was made
   // by Brian Schnieder on 10-24-02.

   // Note that the relative humidities are in ascending order
   // This is required in order for the iteration routine to work correctly

   // If a new value is added, please insert it in the correct ascending location

   // Added to keep relative humidity array bounded by valid values
   dataRelHumidity.nRelativeHumidity = 0;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 10;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 20;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 30;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 40;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 50;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 60;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 70;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 80;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 90;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   dataRelHumidity.nRelativeHumidity = 95;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );

   // Added to keep relative humidity array bounded by valid values
   dataRelHumidity.nRelativeHumidity = 100;
   dataRelHumidity.dIndex = 1.0;
   m_adataRelHumidity.Add( dataRelHumidity );
}

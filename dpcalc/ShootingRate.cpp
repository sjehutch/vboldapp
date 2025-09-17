// ShootingRate.cpp: implementation of the DShootingRate class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "GlobalEnums.h"
#include <math.h>
#include "ShootingRate.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DShootingRate::DShootingRate()
{
}

DShootingRate::~DShootingRate()
{
}


BOOL DShootingRate::GetShootingRate( LPSHOOTINGRATE_DATA lpdataRate )
{
   BOOL bReturn = TRUE;
   double dTimeToEmpty = 0;

   double dRate = GetShootingRate( lpdataRate->dCylinderTemp, 
                                   lpdataRate->dOneFourthLength, 
                                   lpdataRate->dOneEightLength );

   if( dRate > MAX_SHOOTING_RATE )
   {
      dRate = MAX_SHOOTING_RATE;
   }

   if( dRate < MIN_SHOOTING_RATE )
   {
      dRate = MIN_SHOOTING_RATE;
      lpdataRate->strError.LoadString( IDS_WARNING_SHOOTING_RATE_LOW );
   }

   lpdataRate->dShootingRate = dRate;


   // Jett Gamboa 1/14/2005 - added call to compute for the time to empty
   // cylinder
   dTimeToEmpty = GetTimeToEmptyCylinder(dRate,
                                         lpdataRate->dNumCylinders);

   lpdataRate->dTimeToEmptyCylinder = dTimeToEmpty;

   return bReturn;
}


//------------------------------------------------------------------------------
// Determine the shooting rate for a shooting line. Equation below:
//
//                 /---------------------------------------------------------------------
//                / {[ 1469.83 - 6.2374T ][ 10^{4.597575 - [1026.558 / (T + 280.4394)]} ]
// SR = 0.021d^2 /  ---------------------------------------------------------------------
//             \/               /---
//                  {1.5 + [ L / d  (63.34/d - 0.886) ]}
//                           \/
//                            
// SR = Shooting rate, kg/min
// T  = Cylinder temperature, C
// L  = Hose length, M
// d  = Hose inside diameter, mm
// 
// JNM 04/01/03 - Fumiguide Enhancements 2003 - The equation above is no longer used.
//					There would be 3 new equations which would correspond to the 
//					3 Inside Diameter options. The equation to be used will depend
//					on the Inside Diameter:
// 3.175 mm ID line:
// SF flow rate, kg/min = 1.007544 + -0.031001(line length in meters) + 0.056192(temperature in C)
//
// 4.37 mm ID line:
// SF flow rate, kg/min = 2.089543 + -0.095787(line length in meters) + 0.130293(temperature in C)
//
// 6.35 mm ID line:
// SF flow rate, kg/min = 2.63018 + -0.010481(line length in meters) + 0.131075(temperature in C)
//
// MME 05/29/03 - Fumiguide Enhancements 2003 - changed SF Flow Rate formula for 3.175 mm line
// 3/175 mm ID line:
// SF flow rate, kg/min = 1.420949  + (-0.04959 ) * Length of tubing in meters + 0.05704  * temperature (in Fahrenheit)

//------------------------------------------------------------------------------

// CG 09/02/03 - SIMS # 2103035 :  Need to change the introduction rate formula 
//								   for the 6.35 mm (or 1/4") line diamater ONLY.
// METRIC UNITS:
// IF the LENGTH of the 6.35 mm hose inside diameter is between 1 meters and 90 meters:
//		Flow rate in kg/min = 4.873139 + (-0.04652 * line length in meters) + (0.078277 * temperature in degrees C)
// IF the LENGTH of the 6.35 mm hose inside diameter is between 91 meters and 150 meters:
//		Flow rate in kg/min = -1.10071 + (0.014975 * line length in meters) + (0.126524 * temperature in degrees C)
// IF the LENGTH of the 6.35 mm hose inside diameter is greater than 151 meters :
//		Flow rate in kg/min = 5.383333 + (-0.03477 * line length in meters) + (0.108163 * temperature in degrees C)
//
// Jett 09/24/2003 formula has been change to only one for 6.35mm diameter line:
//
// SF flow rate, kg/min = 2.531418 + -0.00849 (line length in meters) + 0.070926(temperature in C)
//
// Pat 12/29/2004 formula has been changed 
//------------------------------------------------------------------------------

double DShootingRate::GetShootingRate( const double &dCylinderTemp, 
                                       const double &dOneFourthLength,
                                       const double &dOneEightLength )
{
	//Pat 12/29/2004 - Removed these and use the new formula
/*	if(dHoseDiameter == 3.175)
		return  1.420949 - 0.04959 * dHoseLength + 0.05704 * dCylinderTemp;
	else if(dHoseDiameter == 4.37)
		return 2.089543 - 0.095787 * dHoseLength + 0.130293 * dCylinderTemp;
	else if(dHoseDiameter == 6.35)
		return 2.531418 - 0.00849 * dHoseLength + 0.070926 * dCylinderTemp;
	else
		return 0.0;*/
	return (1 / ((-0.9238 + (14.21 * (1/dCylinderTemp)) + (0.0416 * pow(dOneFourthLength,0.5)) + (0.1889 * pow(dOneEightLength,0.5)) + ((1/dCylinderTemp) - 0.05654) * (pow(dOneEightLength,0.5) - 1.805) * 4.075)));
}


// Jett Gamboa 3/14/2005 -- added function to calculate time to empty cylinder
// 
// formula provided by businsess is:
// Time to empty cylinders, = {22.51 + 57.14 * 1/(initial flow rate in kg/min)} * number of cylinders

double DShootingRate::GetTimeToEmptyCylinder(const double &dInitialFlowRate, const double &dNumCylinders)
{

	return ((22.51 + 57.14 * (1/dInitialFlowRate)) * dNumCylinders);

}

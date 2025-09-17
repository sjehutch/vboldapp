/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Wed Sep 30 14:25:06 2009
 */
/* Compiler settings for C:\Profume\Source\dpcalc\dpcalc.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __dpcalc_h__
#define __dpcalc_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IProFumeCalculator_FWD_DEFINED__
#define __IProFumeCalculator_FWD_DEFINED__
typedef interface IProFumeCalculator IProFumeCalculator;
#endif 	/* __IProFumeCalculator_FWD_DEFINED__ */


#ifndef __IShootingLine_FWD_DEFINED__
#define __IShootingLine_FWD_DEFINED__
typedef interface IShootingLine IShootingLine;
#endif 	/* __IShootingLine_FWD_DEFINED__ */


#ifndef __ProFumeCalculator_FWD_DEFINED__
#define __ProFumeCalculator_FWD_DEFINED__

#ifdef __cplusplus
typedef class ProFumeCalculator ProFumeCalculator;
#else
typedef struct ProFumeCalculator ProFumeCalculator;
#endif /* __cplusplus */

#endif 	/* __ProFumeCalculator_FWD_DEFINED__ */


#ifndef __ShootingLine_FWD_DEFINED__
#define __ShootingLine_FWD_DEFINED__

#ifdef __cplusplus
typedef class ShootingLine ShootingLine;
#else
typedef struct ShootingLine ShootingLine;
#endif /* __cplusplus */

#endif 	/* __ShootingLine_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IProFumeCalculator_INTERFACE_DEFINED__
#define __IProFumeCalculator_INTERFACE_DEFINED__

/* interface IProFumeCalculator */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IProFumeCalculator;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("8D969523-5811-4E9D-B292-9A16997F9C8B")
    IProFumeCalculator : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPestTotal( 
            /* [retval][out] */ int __RPC_FAR *Total) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPestString( 
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *Pest) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPestIndex( 
            /* [in] */ BSTR __RPC_FAR *Pest,
            /* [retval][out] */ int __RPC_FAR *Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetDosePest( 
            /* [in] */ int PestIndex,
            /* [in] */ int FumigationType,
            /* [in] */ int LifeStage,
            /* [in] */ int TreatmentType,
            /* [in] */ int Commodity,
            /* [in] */ double HLT,
            /* [in] */ double Temp,
            /* [in] */ double TimeExposure,
            /* [in] */ double Volume,
            /* [in] */ double LoadFactor,
            /* [in] */ int IsComplete,
            /* [in] */ int ICounter,
            /* [out] */ BSTR __RPC_FAR *Error) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetProfumeRequired( 
            /* [out] */ double __RPC_FAR *ProfumeReq,
            /* [out] */ double __RPC_FAR *Cylinders,
            /* [out] */ double __RPC_FAR *Co,
            /* [out] */ double __RPC_FAR *CT,
            /* [out] */ double __RPC_FAR *Dose,
            /* [out] */ BSTR __RPC_FAR *Error) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ClearDosePest( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetMonitorPointStatus( 
            /* [in] */ double OriginalHLT,
            /* [out] */ double __RPC_FAR *AchievedCT,
            /* [out] */ double __RPC_FAR *AchievedDose,
            /* [out] */ double __RPC_FAR *ActualOHLT,
            /* [out] */ double __RPC_FAR *ProjectedCT,
            /* [out] */ double __RPC_FAR *ProjectedFumigationTime,
            /* [out] */ double __RPC_FAR *FutureTargetConcentration,
            /* [out] */ double __RPC_FAR *FutureConcentration,
            /* [out] */ double __RPC_FAR *ElapsedTime,
            /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
            /* [out] */ DATE __RPC_FAR *HLTDateTime,
            /* [out] */ BSTR __RPC_FAR *Msg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetConcentration( 
            /* [in] */ double Concentration,
            /* [in] */ DATE Date,
            /* [out] */ BSTR __RPC_FAR *Msg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ClearConcentrations( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetVersion( 
            /* [retval][out] */ BSTR __RPC_FAR *Version) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPt2PtHlt( 
            /* [out] */ double __RPC_FAR *LastPeakHLT,
            /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
            /* [out] */ DATE __RPC_FAR *HLTDateTime,
            /* [out] */ BSTR __RPC_FAR *Msg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetAreaMonitorPoint( 
            /* [in] */ double AchievedConcentrationTime,
            /* [in] */ double AchievedDose,
            /* [in] */ double ElapsedTime,
            /* [in] */ double FutureTargetConcentration,
            /* [in] */ double FutureConcentration,
            /* [in] */ double Hlt,
            /* [in] */ double ProjectedConcentrationTime,
            /* [in] */ double ProjectedTime,
            /* [out] */ BSTR __RPC_FAR *Msg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ClearAreaMonitorPoints( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetAreaStatus( 
            /* [in] */ double TargetConcentrationTime,
            /* [out] */ double __RPC_FAR *AchievedConcentrationTime,
            /* [out] */ double __RPC_FAR *AchievedDose,
            /* [out] */ double __RPC_FAR *Hlt,
            /* [out] */ double __RPC_FAR *ProjectedConcentrationTime,
            /* [out] */ double __RPC_FAR *ProjectedTime,
            /* [out] */ double __RPC_FAR *MeanNTC,
            /* [out] */ int __RPC_FAR *Status,
            /* [out] */ double __RPC_FAR *AddProfume,
            /* [out] */ BSTR __RPC_FAR *Msg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetCommodityIndex( 
            /* [in] */ BSTR __RPC_FAR *Commodity,
            /* [retval][out] */ int __RPC_FAR *Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetCommodityString( 
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *Commodity) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetCommodityTotal( 
            /* [retval][out] */ int __RPC_FAR *Total) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetLifeStageTotal( 
            /* [retval][out] */ int __RPC_FAR *Total) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetLifeStageIndex( 
            /* [in] */ BSTR __RPC_FAR *LifeStage,
            /* [retval][out] */ int __RPC_FAR *Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetLifeStageString( 
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *LifeStage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFumigationTypeTotal( 
            /* [retval][out] */ int __RPC_FAR *Total) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFumigationTypeIndex( 
            /* [in] */ BSTR __RPC_FAR *FumigationType,
            /* [retval][out] */ int __RPC_FAR *Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFumigationTypeString( 
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *FumigationType) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetTreatmentTypeTotal( 
            /* [retval][out] */ int __RPC_FAR *Total) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetTreatmentTypeIndex( 
            /* [in] */ BSTR __RPC_FAR *TreatmentType,
            /* [retval][out] */ int __RPC_FAR *Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetTreatmentTypeString( 
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *TreatmentType) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IProFumeCalculatorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IProFumeCalculator __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IProFumeCalculator __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPestTotal )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *Total);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPestString )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *Pest);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPestIndex )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ BSTR __RPC_FAR *Pest,
            /* [retval][out] */ int __RPC_FAR *Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetDosePest )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int PestIndex,
            /* [in] */ int FumigationType,
            /* [in] */ int LifeStage,
            /* [in] */ int TreatmentType,
            /* [in] */ int Commodity,
            /* [in] */ double HLT,
            /* [in] */ double Temp,
            /* [in] */ double TimeExposure,
            /* [in] */ double Volume,
            /* [in] */ double LoadFactor,
            /* [in] */ int IsComplete,
            /* [in] */ int ICounter,
            /* [out] */ BSTR __RPC_FAR *Error);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetProfumeRequired )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [out] */ double __RPC_FAR *ProfumeReq,
            /* [out] */ double __RPC_FAR *Cylinders,
            /* [out] */ double __RPC_FAR *Co,
            /* [out] */ double __RPC_FAR *CT,
            /* [out] */ double __RPC_FAR *Dose,
            /* [out] */ BSTR __RPC_FAR *Error);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ClearDosePest )( 
            IProFumeCalculator __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetMonitorPointStatus )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ double OriginalHLT,
            /* [out] */ double __RPC_FAR *AchievedCT,
            /* [out] */ double __RPC_FAR *AchievedDose,
            /* [out] */ double __RPC_FAR *ActualOHLT,
            /* [out] */ double __RPC_FAR *ProjectedCT,
            /* [out] */ double __RPC_FAR *ProjectedFumigationTime,
            /* [out] */ double __RPC_FAR *FutureTargetConcentration,
            /* [out] */ double __RPC_FAR *FutureConcentration,
            /* [out] */ double __RPC_FAR *ElapsedTime,
            /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
            /* [out] */ DATE __RPC_FAR *HLTDateTime,
            /* [out] */ BSTR __RPC_FAR *Msg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetConcentration )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ double Concentration,
            /* [in] */ DATE Date,
            /* [out] */ BSTR __RPC_FAR *Msg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ClearConcentrations )( 
            IProFumeCalculator __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetVersion )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *Version);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPt2PtHlt )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [out] */ double __RPC_FAR *LastPeakHLT,
            /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
            /* [out] */ DATE __RPC_FAR *HLTDateTime,
            /* [out] */ BSTR __RPC_FAR *Msg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetAreaMonitorPoint )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ double AchievedConcentrationTime,
            /* [in] */ double AchievedDose,
            /* [in] */ double ElapsedTime,
            /* [in] */ double FutureTargetConcentration,
            /* [in] */ double FutureConcentration,
            /* [in] */ double Hlt,
            /* [in] */ double ProjectedConcentrationTime,
            /* [in] */ double ProjectedTime,
            /* [out] */ BSTR __RPC_FAR *Msg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ClearAreaMonitorPoints )( 
            IProFumeCalculator __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetAreaStatus )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ double TargetConcentrationTime,
            /* [out] */ double __RPC_FAR *AchievedConcentrationTime,
            /* [out] */ double __RPC_FAR *AchievedDose,
            /* [out] */ double __RPC_FAR *Hlt,
            /* [out] */ double __RPC_FAR *ProjectedConcentrationTime,
            /* [out] */ double __RPC_FAR *ProjectedTime,
            /* [out] */ double __RPC_FAR *MeanNTC,
            /* [out] */ int __RPC_FAR *Status,
            /* [out] */ double __RPC_FAR *AddProfume,
            /* [out] */ BSTR __RPC_FAR *Msg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCommodityIndex )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ BSTR __RPC_FAR *Commodity,
            /* [retval][out] */ int __RPC_FAR *Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCommodityString )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *Commodity);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCommodityTotal )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *Total);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLifeStageTotal )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *Total);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLifeStageIndex )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ BSTR __RPC_FAR *LifeStage,
            /* [retval][out] */ int __RPC_FAR *Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLifeStageString )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *LifeStage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFumigationTypeTotal )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *Total);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFumigationTypeIndex )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ BSTR __RPC_FAR *FumigationType,
            /* [retval][out] */ int __RPC_FAR *Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFumigationTypeString )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *FumigationType);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTreatmentTypeTotal )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *Total);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTreatmentTypeIndex )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ BSTR __RPC_FAR *TreatmentType,
            /* [retval][out] */ int __RPC_FAR *Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTreatmentTypeString )( 
            IProFumeCalculator __RPC_FAR * This,
            /* [in] */ int Index,
            /* [retval][out] */ BSTR __RPC_FAR *TreatmentType);
        
        END_INTERFACE
    } IProFumeCalculatorVtbl;

    interface IProFumeCalculator
    {
        CONST_VTBL struct IProFumeCalculatorVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IProFumeCalculator_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IProFumeCalculator_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IProFumeCalculator_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IProFumeCalculator_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IProFumeCalculator_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IProFumeCalculator_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IProFumeCalculator_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IProFumeCalculator_GetPestTotal(This,Total)	\
    (This)->lpVtbl -> GetPestTotal(This,Total)

#define IProFumeCalculator_GetPestString(This,Index,Pest)	\
    (This)->lpVtbl -> GetPestString(This,Index,Pest)

#define IProFumeCalculator_GetPestIndex(This,Pest,Index)	\
    (This)->lpVtbl -> GetPestIndex(This,Pest,Index)

#define IProFumeCalculator_SetDosePest(This,PestIndex,FumigationType,LifeStage,TreatmentType,Commodity,HLT,Temp,TimeExposure,Volume,LoadFactor,IsComplete,ICounter,Error)	\
    (This)->lpVtbl -> SetDosePest(This,PestIndex,FumigationType,LifeStage,TreatmentType,Commodity,HLT,Temp,TimeExposure,Volume,LoadFactor,IsComplete,ICounter,Error)

#define IProFumeCalculator_GetProfumeRequired(This,ProfumeReq,Cylinders,Co,CT,Dose,Error)	\
    (This)->lpVtbl -> GetProfumeRequired(This,ProfumeReq,Cylinders,Co,CT,Dose,Error)

#define IProFumeCalculator_ClearDosePest(This)	\
    (This)->lpVtbl -> ClearDosePest(This)

#define IProFumeCalculator_GetMonitorPointStatus(This,OriginalHLT,AchievedCT,AchievedDose,ActualOHLT,ProjectedCT,ProjectedFumigationTime,FutureTargetConcentration,FutureConcentration,ElapsedTime,LastPeakDateTime,HLTDateTime,Msg)	\
    (This)->lpVtbl -> GetMonitorPointStatus(This,OriginalHLT,AchievedCT,AchievedDose,ActualOHLT,ProjectedCT,ProjectedFumigationTime,FutureTargetConcentration,FutureConcentration,ElapsedTime,LastPeakDateTime,HLTDateTime,Msg)

#define IProFumeCalculator_SetConcentration(This,Concentration,Date,Msg)	\
    (This)->lpVtbl -> SetConcentration(This,Concentration,Date,Msg)

#define IProFumeCalculator_ClearConcentrations(This)	\
    (This)->lpVtbl -> ClearConcentrations(This)

#define IProFumeCalculator_GetVersion(This,Version)	\
    (This)->lpVtbl -> GetVersion(This,Version)

#define IProFumeCalculator_GetPt2PtHlt(This,LastPeakHLT,LastPeakDateTime,HLTDateTime,Msg)	\
    (This)->lpVtbl -> GetPt2PtHlt(This,LastPeakHLT,LastPeakDateTime,HLTDateTime,Msg)

#define IProFumeCalculator_SetAreaMonitorPoint(This,AchievedConcentrationTime,AchievedDose,ElapsedTime,FutureTargetConcentration,FutureConcentration,Hlt,ProjectedConcentrationTime,ProjectedTime,Msg)	\
    (This)->lpVtbl -> SetAreaMonitorPoint(This,AchievedConcentrationTime,AchievedDose,ElapsedTime,FutureTargetConcentration,FutureConcentration,Hlt,ProjectedConcentrationTime,ProjectedTime,Msg)

#define IProFumeCalculator_ClearAreaMonitorPoints(This)	\
    (This)->lpVtbl -> ClearAreaMonitorPoints(This)

#define IProFumeCalculator_GetAreaStatus(This,TargetConcentrationTime,AchievedConcentrationTime,AchievedDose,Hlt,ProjectedConcentrationTime,ProjectedTime,MeanNTC,Status,AddProfume,Msg)	\
    (This)->lpVtbl -> GetAreaStatus(This,TargetConcentrationTime,AchievedConcentrationTime,AchievedDose,Hlt,ProjectedConcentrationTime,ProjectedTime,MeanNTC,Status,AddProfume,Msg)

#define IProFumeCalculator_GetCommodityIndex(This,Commodity,Index)	\
    (This)->lpVtbl -> GetCommodityIndex(This,Commodity,Index)

#define IProFumeCalculator_GetCommodityString(This,Index,Commodity)	\
    (This)->lpVtbl -> GetCommodityString(This,Index,Commodity)

#define IProFumeCalculator_GetCommodityTotal(This,Total)	\
    (This)->lpVtbl -> GetCommodityTotal(This,Total)

#define IProFumeCalculator_GetLifeStageTotal(This,Total)	\
    (This)->lpVtbl -> GetLifeStageTotal(This,Total)

#define IProFumeCalculator_GetLifeStageIndex(This,LifeStage,Index)	\
    (This)->lpVtbl -> GetLifeStageIndex(This,LifeStage,Index)

#define IProFumeCalculator_GetLifeStageString(This,Index,LifeStage)	\
    (This)->lpVtbl -> GetLifeStageString(This,Index,LifeStage)

#define IProFumeCalculator_GetFumigationTypeTotal(This,Total)	\
    (This)->lpVtbl -> GetFumigationTypeTotal(This,Total)

#define IProFumeCalculator_GetFumigationTypeIndex(This,FumigationType,Index)	\
    (This)->lpVtbl -> GetFumigationTypeIndex(This,FumigationType,Index)

#define IProFumeCalculator_GetFumigationTypeString(This,Index,FumigationType)	\
    (This)->lpVtbl -> GetFumigationTypeString(This,Index,FumigationType)

#define IProFumeCalculator_GetTreatmentTypeTotal(This,Total)	\
    (This)->lpVtbl -> GetTreatmentTypeTotal(This,Total)

#define IProFumeCalculator_GetTreatmentTypeIndex(This,TreatmentType,Index)	\
    (This)->lpVtbl -> GetTreatmentTypeIndex(This,TreatmentType,Index)

#define IProFumeCalculator_GetTreatmentTypeString(This,Index,TreatmentType)	\
    (This)->lpVtbl -> GetTreatmentTypeString(This,Index,TreatmentType)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetPestTotal_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *Total);


void __RPC_STUB IProFumeCalculator_GetPestTotal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetPestString_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int Index,
    /* [retval][out] */ BSTR __RPC_FAR *Pest);


void __RPC_STUB IProFumeCalculator_GetPestString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetPestIndex_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ BSTR __RPC_FAR *Pest,
    /* [retval][out] */ int __RPC_FAR *Index);


void __RPC_STUB IProFumeCalculator_GetPestIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_SetDosePest_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int PestIndex,
    /* [in] */ int FumigationType,
    /* [in] */ int LifeStage,
    /* [in] */ int TreatmentType,
    /* [in] */ int Commodity,
    /* [in] */ double HLT,
    /* [in] */ double Temp,
    /* [in] */ double TimeExposure,
    /* [in] */ double Volume,
    /* [in] */ double LoadFactor,
    /* [in] */ int IsComplete,
    /* [in] */ int ICounter,
    /* [out] */ BSTR __RPC_FAR *Error);


void __RPC_STUB IProFumeCalculator_SetDosePest_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetProfumeRequired_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [out] */ double __RPC_FAR *ProfumeReq,
    /* [out] */ double __RPC_FAR *Cylinders,
    /* [out] */ double __RPC_FAR *Co,
    /* [out] */ double __RPC_FAR *CT,
    /* [out] */ double __RPC_FAR *Dose,
    /* [out] */ BSTR __RPC_FAR *Error);


void __RPC_STUB IProFumeCalculator_GetProfumeRequired_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_ClearDosePest_Proxy( 
    IProFumeCalculator __RPC_FAR * This);


void __RPC_STUB IProFumeCalculator_ClearDosePest_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetMonitorPointStatus_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ double OriginalHLT,
    /* [out] */ double __RPC_FAR *AchievedCT,
    /* [out] */ double __RPC_FAR *AchievedDose,
    /* [out] */ double __RPC_FAR *ActualOHLT,
    /* [out] */ double __RPC_FAR *ProjectedCT,
    /* [out] */ double __RPC_FAR *ProjectedFumigationTime,
    /* [out] */ double __RPC_FAR *FutureTargetConcentration,
    /* [out] */ double __RPC_FAR *FutureConcentration,
    /* [out] */ double __RPC_FAR *ElapsedTime,
    /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
    /* [out] */ DATE __RPC_FAR *HLTDateTime,
    /* [out] */ BSTR __RPC_FAR *Msg);


void __RPC_STUB IProFumeCalculator_GetMonitorPointStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_SetConcentration_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ double Concentration,
    /* [in] */ DATE Date,
    /* [out] */ BSTR __RPC_FAR *Msg);


void __RPC_STUB IProFumeCalculator_SetConcentration_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_ClearConcentrations_Proxy( 
    IProFumeCalculator __RPC_FAR * This);


void __RPC_STUB IProFumeCalculator_ClearConcentrations_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetVersion_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *Version);


void __RPC_STUB IProFumeCalculator_GetVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetPt2PtHlt_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [out] */ double __RPC_FAR *LastPeakHLT,
    /* [out] */ DATE __RPC_FAR *LastPeakDateTime,
    /* [out] */ DATE __RPC_FAR *HLTDateTime,
    /* [out] */ BSTR __RPC_FAR *Msg);


void __RPC_STUB IProFumeCalculator_GetPt2PtHlt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_SetAreaMonitorPoint_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ double AchievedConcentrationTime,
    /* [in] */ double AchievedDose,
    /* [in] */ double ElapsedTime,
    /* [in] */ double FutureTargetConcentration,
    /* [in] */ double FutureConcentration,
    /* [in] */ double Hlt,
    /* [in] */ double ProjectedConcentrationTime,
    /* [in] */ double ProjectedTime,
    /* [out] */ BSTR __RPC_FAR *Msg);


void __RPC_STUB IProFumeCalculator_SetAreaMonitorPoint_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_ClearAreaMonitorPoints_Proxy( 
    IProFumeCalculator __RPC_FAR * This);


void __RPC_STUB IProFumeCalculator_ClearAreaMonitorPoints_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetAreaStatus_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ double TargetConcentrationTime,
    /* [out] */ double __RPC_FAR *AchievedConcentrationTime,
    /* [out] */ double __RPC_FAR *AchievedDose,
    /* [out] */ double __RPC_FAR *Hlt,
    /* [out] */ double __RPC_FAR *ProjectedConcentrationTime,
    /* [out] */ double __RPC_FAR *ProjectedTime,
    /* [out] */ double __RPC_FAR *MeanNTC,
    /* [out] */ int __RPC_FAR *Status,
    /* [out] */ double __RPC_FAR *AddProfume,
    /* [out] */ BSTR __RPC_FAR *Msg);


void __RPC_STUB IProFumeCalculator_GetAreaStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetCommodityIndex_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ BSTR __RPC_FAR *Commodity,
    /* [retval][out] */ int __RPC_FAR *Index);


void __RPC_STUB IProFumeCalculator_GetCommodityIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetCommodityString_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int Index,
    /* [retval][out] */ BSTR __RPC_FAR *Commodity);


void __RPC_STUB IProFumeCalculator_GetCommodityString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetCommodityTotal_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *Total);


void __RPC_STUB IProFumeCalculator_GetCommodityTotal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetLifeStageTotal_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *Total);


void __RPC_STUB IProFumeCalculator_GetLifeStageTotal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetLifeStageIndex_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ BSTR __RPC_FAR *LifeStage,
    /* [retval][out] */ int __RPC_FAR *Index);


void __RPC_STUB IProFumeCalculator_GetLifeStageIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetLifeStageString_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int Index,
    /* [retval][out] */ BSTR __RPC_FAR *LifeStage);


void __RPC_STUB IProFumeCalculator_GetLifeStageString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetFumigationTypeTotal_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *Total);


void __RPC_STUB IProFumeCalculator_GetFumigationTypeTotal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetFumigationTypeIndex_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ BSTR __RPC_FAR *FumigationType,
    /* [retval][out] */ int __RPC_FAR *Index);


void __RPC_STUB IProFumeCalculator_GetFumigationTypeIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetFumigationTypeString_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int Index,
    /* [retval][out] */ BSTR __RPC_FAR *FumigationType);


void __RPC_STUB IProFumeCalculator_GetFumigationTypeString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetTreatmentTypeTotal_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *Total);


void __RPC_STUB IProFumeCalculator_GetTreatmentTypeTotal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetTreatmentTypeIndex_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ BSTR __RPC_FAR *TreatmentType,
    /* [retval][out] */ int __RPC_FAR *Index);


void __RPC_STUB IProFumeCalculator_GetTreatmentTypeIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProFumeCalculator_GetTreatmentTypeString_Proxy( 
    IProFumeCalculator __RPC_FAR * This,
    /* [in] */ int Index,
    /* [retval][out] */ BSTR __RPC_FAR *TreatmentType);


void __RPC_STUB IProFumeCalculator_GetTreatmentTypeString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IProFumeCalculator_INTERFACE_DEFINED__ */


#ifndef __IShootingLine_INTERFACE_DEFINED__
#define __IShootingLine_INTERFACE_DEFINED__

/* interface IShootingLine */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IShootingLine;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("B11CCE60-3C64-4a1c-9650-63AC20D54E5D")
    IShootingLine : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetShootingRate( 
            /* [in] */ double CylinderTemp,
            /* [in] */ double OneFourthLength,
            /* [in] */ double OneEightLength,
            /* [in] */ double NumCylinders,
            /* [in] */ double PressureChosen,
            /* [out] */ double __RPC_FAR *ShootingRate,
            /* [out] */ double __RPC_FAR *TimeToEmptyCylinder,
            /* [out] */ BSTR __RPC_FAR *ErrMsg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetMaxShootingRate( 
            /* [in] */ double FanCapacity,
            /* [in] */ int RelativeHumidity,
            /* [out] */ double __RPC_FAR *MaxShootingRate,
            /* [out] */ BSTR __RPC_FAR *ErrMsg) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IShootingLineVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IShootingLine __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IShootingLine __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IShootingLine __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetShootingRate )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ double CylinderTemp,
            /* [in] */ double OneFourthLength,
            /* [in] */ double OneEightLength,
            /* [in] */ double NumCylinders,
            /* [in] */ double PressureChosen,
            /* [out] */ double __RPC_FAR *ShootingRate,
            /* [out] */ double __RPC_FAR *TimeToEmptyCylinder,
            /* [out] */ BSTR __RPC_FAR *ErrMsg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetMaxShootingRate )( 
            IShootingLine __RPC_FAR * This,
            /* [in] */ double FanCapacity,
            /* [in] */ int RelativeHumidity,
            /* [out] */ double __RPC_FAR *MaxShootingRate,
            /* [out] */ BSTR __RPC_FAR *ErrMsg);
        
        END_INTERFACE
    } IShootingLineVtbl;

    interface IShootingLine
    {
        CONST_VTBL struct IShootingLineVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IShootingLine_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IShootingLine_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IShootingLine_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IShootingLine_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IShootingLine_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IShootingLine_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IShootingLine_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IShootingLine_GetShootingRate(This,CylinderTemp,OneFourthLength,OneEightLength,NumCylinders,PressureChosen,ShootingRate,TimeToEmptyCylinder,ErrMsg)	\
    (This)->lpVtbl -> GetShootingRate(This,CylinderTemp,OneFourthLength,OneEightLength,NumCylinders,PressureChosen,ShootingRate,TimeToEmptyCylinder,ErrMsg)

#define IShootingLine_GetMaxShootingRate(This,FanCapacity,RelativeHumidity,MaxShootingRate,ErrMsg)	\
    (This)->lpVtbl -> GetMaxShootingRate(This,FanCapacity,RelativeHumidity,MaxShootingRate,ErrMsg)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IShootingLine_GetShootingRate_Proxy( 
    IShootingLine __RPC_FAR * This,
    /* [in] */ double CylinderTemp,
    /* [in] */ double OneFourthLength,
    /* [in] */ double OneEightLength,
    /* [in] */ double NumCylinders,
    /* [in] */ double PressureChosen,
    /* [out] */ double __RPC_FAR *ShootingRate,
    /* [out] */ double __RPC_FAR *TimeToEmptyCylinder,
    /* [out] */ BSTR __RPC_FAR *ErrMsg);


void __RPC_STUB IShootingLine_GetShootingRate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IShootingLine_GetMaxShootingRate_Proxy( 
    IShootingLine __RPC_FAR * This,
    /* [in] */ double FanCapacity,
    /* [in] */ int RelativeHumidity,
    /* [out] */ double __RPC_FAR *MaxShootingRate,
    /* [out] */ BSTR __RPC_FAR *ErrMsg);


void __RPC_STUB IShootingLine_GetMaxShootingRate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IShootingLine_INTERFACE_DEFINED__ */



#ifndef __DPCALCLib_LIBRARY_DEFINED__
#define __DPCALCLib_LIBRARY_DEFINED__

/* library DPCALCLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_DPCALCLib;

EXTERN_C const CLSID CLSID_ProFumeCalculator;

#ifdef __cplusplus

class DECLSPEC_UUID("45DD8AD9-D339-4A0B-B6FE-260433177DD6")
ProFumeCalculator;
#endif

EXTERN_C const CLSID CLSID_ShootingLine;

#ifdef __cplusplus

class DECLSPEC_UUID("A11B8F6C-BBA7-46ba-9250-358350DF9ABA")
ShootingLine;
#endif
#endif /* __DPCALCLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif

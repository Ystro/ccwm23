
using Progress.Lang.*.

block-level on error undo, throw.

class MISSINGNAME: 
  
  method public character cSecondStep():
    if not lCompletedStepsBefore(2) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */
        
    prpoCurrentDocument = prpoWord:SetUpPHDocument(oBSC:GeneratePHNubmer()).
    
    prpoConsultant:writeSpecification(oDocument).
    
    BBCOutlookSvc:prpoInstance:sendMail(
      substitute ('Pflichtenheft &1':U, oDocument:prpcName),
      prpoConsultant:prpcEMail,
      prpoCustomer:prpcEmail,
      oDocument
      ).
  
    return 'PH ????????? und ?????????':U.
  end method.
  
  method public character cThirdStep():
    if not lCompletedStepsBefore(3) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */

    prpoEngineer = BBCITCPlanningSvc:prpoInstance:oGetNextFreeEngineer(prpoCustomer:prpcRegion).
    prpoEngineer:writeTechDetails(prpoDocument).
    
    prpoCurrentSpecification = prpoExcel:SetUpSpecification(prpoDocument).
    prpoEngineer:calculateEffort(prpoCurrentSpecification, prpoDocument).
    
    prpoModificationPlan = BBCProjectsSvc:prpoInstance:sendOffer(prpoCurrentSpecification, prpoCustomer).
    
  
    return '?????????? ????????????? ??? Angebot':U.
  end method.
  
  method public character cFourthStep():
    if not lCompletedStepsBefore(4) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */

    prpoPlanningInformations = BBCITCPlanningSvc:prpoInstance:planModification(
        prpoModificationPlan, /* from here we get the current modification */
        prpoEngineer /* preferred developer should be whoever has done the specification */
      ). 
      
    prpoPlanningInformations:prpoTargetEngineer:approveProjectPlan(oProjectPlan).
    BBCProjectsSvc:prpoInstance:sendConfirmationWithDeliveryDate(prpoPlanningInformations, prpoCustomer).
    
    return 'Beauftragung ??? ??????????':U.
  end method.
  
  method public character cFifthStep():
    if not lCompletedStepsBefore(5) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */
    
    if not prpoDeliveryPackage then
      prpoDeliveryPackage = prpoPlanningInformations:prpoTargetEngineer:oImplementSolution().
    else 
      prpoDeliveryPackage = prpoPlanningInformations:prpoTargetEngineer:oImplementFixes().
    
    BBCTransferChangeTeamSvc:prpoInstance:deliverPackageToTest(prpoDeliveryPackage).
    
    lDevTestSuccess = prpoPlanningInformations:prpoTargetEngineer:lTestImplementation(prpoDeliveryPackage).
    
    if (lDevTestSuccess)
      return '?????????, ???????????? und ??????????????':U.
    else
      return cFifthStep().
  end method.
  
  method public character cSixthStep():
    if not lCompletedStepsBefore(6) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */

    oFinalReview = prpoConsultant:oTestAndReview().
    if not oFinalReview:prplSuccess then do:
      BBCProjectsSvc:prpoInstance:generateReworkTicket(prpoModificationPlan).
      cRework().
      return 
        adm.method.cls.DMCMessageSvc:prpoInstance:showError
          ('bpbcs0003':U). /* Consultant Tests not ok */
    end.
      
    BBCOutlookSvc:prpoInstance:sendMail(
      substitute('&1 für Sie testbereit!':U, oFinalReview:prpcName),
      prpoConsultant:prpcEMail,
      prpoCustomer:prpcEmail,
      oFinalReview
    ).
  
    return '??????????? und ????????? ??? Betriebsbereitschaft':U.
  end method.
  
  method public character cSeventhStep():
    if not lCompletedStepsBefore(7) then 
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0002':U). /* not all steps before finished */

    prpoCustomerFeedback = prpoCustomer:testAndProvideFeedback().
    
    if prpoCustomerFeedback:prplTestSucess then
    do:
      BBCTransferChangeTeamSvc:prpoInstance:deliverPackageToProduction(prpoDeliveryPackage).
      prpoModificationPlan:closeModificationTask().
      return 'Kund******':U.
    end.
    else do:
      BBCProjectsSvc:prpoInstance:generateReworkTicket(prpoModificationPlan).
      cRework().
      adm.method.cls.DMCMessageSvc:prpoInstance:showError
        ('bpbcs0004':U). /* Customer Tests not ok */
    end.
    
  end method.
  

end class.
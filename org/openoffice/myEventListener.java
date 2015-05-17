
package org.openoffice;

import com.sun.star.view.XSelectionChangeListener;


class myEventListener implements XSelectionChangeListener{

    private myAddOn target = null;

    public myEventListener(myAddOn addOn) {
        target = addOn;
    }

    
    public void disposing(com.sun.star.lang.EventObject event) {
       //System.out.println("desposing myEventListener");
    }


    public void selectionChanged(com.sun.star.lang.EventObject event) {
        try{
        //System.out.println("selection changed ag");
            target.setCurrentCellValue(event, 1);
        }catch(Exception e){
            System.out.println("problem with selectionChanged() " +
                e.getMessage());
            e.printStackTrace();
        }
    }

}

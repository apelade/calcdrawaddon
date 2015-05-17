package org.openoffice;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.awt.Size;
import com.sun.star.frame.XDispatch;
import com.sun.star.lang.EventObject;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.view.XSelectionSupplier;
import java.awt.Point;
import java.util.LinkedHashSet;
import java.util.Set;

public final class myAddOn extends WeakBase
        implements com.sun.star.lang.XInitialization,
        com.sun.star.frame.XDispatch,
        com.sun.star.lang.XServiceInfo,
        com.sun.star.frame.XDispatchProvider {

    private final XComponentContext m_xContext;
    private com.sun.star.frame.XFrame m_xFrame;
    private static final String m_implementationName = myAddOn.class.getName();
    private static final String[] m_serviceNames = {
        "com.sun.star.frame.ProtocolHandler"};
    
////////////////////////////////////////////////////////////////////////////////
//////           User changes between here and the next marker              ////
////////////////////////////////////////////////////////////////////////////////
    
    private final myEventListener eventListener;
    private static final String CELL_BACK_COLOR = "CellBackColor";
    private boolean listen = false;
    //@todo: make connect another option menu item of this addOn
    //For now though, static final CAPS to show people this is the switch point.
    private static final boolean CONNECT = true;
    private boolean dirtyFlag = false; // dirty flag checked in getSpots()
    private int startCol = -1;
    private int startRow = -1;
    private int endCol = -1;
    private int endRow = -1;
    private int currentColumn = -1;
    private int currentRow = -1;
    private int cellWidth = -1;
    private int cellHeight = -1;

    public myAddOn(XComponentContext context) {
        m_xContext = context;
        eventListener = new myEventListener(this);
    }

    // com.sun.star.frame.XDispatch:
    public void dispatch(com.sun.star.util.URL aURL,
            com.sun.star.beans.PropertyValue[] aArguments) {
        if (aURL.Protocol.compareTo("org.openoffice.myaddon:") == 0) {
            if (aURL.Path.compareTo("FingerPaint") == 0) {
                System.out.println("dispatch() listen set to "
                        + toggleListen());
            }
        }
        initCellSize();
        resetFirstTouch();
    }

    // called from the selection change listener
    // cellValue used for setting xCell.setValue()
    protected void setCurrentCellValue(EventObject event, double cellValue) {
        CellRangeAddress cra = getCurrentCellRange(event);
        XModel model = m_xFrame.getController().getModel();
        XSpreadsheetDocument doc =
                (XSpreadsheetDocument) UnoRuntime.queryInterface(
                XSpreadsheetDocument.class, model);
        XSpreadsheet sheet = getSpreadsheetByIndex(doc, cra.Sheet);
        if (CONNECT == true) {
            Point oldPoint = new Point(currentColumn, currentRow);
            updateCurrentPoint(cra); // changes currentColumn and currentRow
            Point newPoint = new Point(currentColumn, currentRow);
            Set<Point> points = getSpots(
                    oldPoint,
                    newPoint,
                    cellWidth,
                    cellHeight);
            // freeze painting?
            for (Point pt : points){
                makeCellChangeAt(sheet, pt.x, pt.y);
            }
            // repaint()
        } else {
            updateCurrentPoint(cra);
            makeCellChangeAt(sheet, currentColumn, currentRow);
        }
    }

    public CellRangeAddress getCurrentCellRange(EventObject event) {
        XModel model = m_xFrame.getController().getModel();
        XCellRangeAddressable xSheetCellAddressable =
                (XCellRangeAddressable) UnoRuntime.queryInterface(
                XCellRangeAddressable.class, model.getCurrentSelection());
        if (xSheetCellAddressable != null) {
            return xSheetCellAddressable.getRangeAddress();
        } else {
            return null;
        }
    }
    
    /** Returns the spreadsheet with the specified index (0-based).
    @param xDocument The XSpreadsheetDocument interface of the document.
    @param nIndex The index of the sheet.
    @return The XSpreadsheet interface of the sheet. */
    public static XSpreadsheet getSpreadsheetByIndex(
            XSpreadsheetDocument xDocument, int nIndex) {
        XSpreadsheets xSheets = xDocument.getSheets();
        XSpreadsheet xSheet = null;
        try {
            XIndexAccess xSheetsIA = (XIndexAccess) UnoRuntime.queryInterface(
                    XIndexAccess.class, xSheets);
            Object sheetObj = xSheetsIA.getByIndex(nIndex);
            xSheet = (XSpreadsheet) UnoRuntime.queryInterface(
                    XSpreadsheet.class, sheetObj);
        } catch (Exception ex) {
            System.err.println("getSpreadsheet problemo");
            System.out.println(ex);
        }
        return xSheet;
    }

    protected void updateCurrentPoint(CellRangeAddress cra) {

        if (cra.StartColumn != startCol) {
            currentColumn = cra.StartColumn;
        } else if (cra.EndColumn != endCol) {
            currentColumn = cra.EndColumn;
        }
        if (cra.StartRow != startRow) {
            currentRow = cra.StartRow;
        } else if (cra.EndRow != endRow) {
            currentRow = cra.EndRow;
        }
        startCol = cra.StartColumn;
        startRow = cra.StartRow;
        endCol = cra.EndColumn;
        endRow = cra.EndRow;
    }


    private void makeCellChangeAt(XSpreadsheet sheet, int col, int row) {
        try {
            XCell cell = sheet.getCellByPosition(col, row);
            // As an alternative, use Calc>Tools>Options>Conditional format
            // on the document and set the background color by value:
            // cell.setValue(cellValue);
            XPropertySet cellProps =
                    (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, cell);
            cellProps.setPropertyValue(
                    myAddOn.CELL_BACK_COLOR, getCurrentColorInteger());
        } catch (UnknownPropertyException ex) {
            System.out.println(ex);
        } catch (PropertyVetoException ex) {
            System.out.println(ex);
        } catch (IllegalArgumentException ex) {
            System.out.println(ex);
        } catch (WrappedTargetException ex) {
            System.out.println(ex);
        } catch (NullPointerException ex) {
            System.out.println("null prob getting cell " + ex);
        } catch (IndexOutOfBoundsException ex) {
            System.out.println("index prob getting cell " + ex);
        }
    }

    public Set<Point> getSpots(
            Point start,
            Point stop,
            int width,
            int height){
        LinkedHashSet spots = new LinkedHashSet();
       //if(start.x == -1 && start.y == -1){
        if(dirtyFlag == false){
            dirtyFlag = true;
            spots.add(stop);
            return spots;
        }
        double xch = stop.x - start.x;
        double ych = stop.y - start.y;
        double numSpots = Math.max(Math.abs(xch), Math.abs(ych));
        if (numSpots == 0) {
            spots.add(start);
            return spots;
        } else {
            double lenx = (xch / numSpots) * width;
            double leny = (ych / numSpots) * height;
            double centerx = (start.x * width) + (0.5 * width);
            double centery = (start.y * height) + (0.5 * height);
            for (int i = 0; i < numSpots; i++) {
                Point gridPt = new Point(
                        (int) centerx / width,
                        (int) centery / height);
                spots.add(gridPt);
                centerx += lenx;
                centery += leny;
            }
        }
        return spots;
    }

    private void initCellSize() {
        XModel model = m_xFrame.getController().getModel();
        XSpreadsheetDocument doc =
                (XSpreadsheetDocument) UnoRuntime.queryInterface(
                XSpreadsheetDocument.class, model);
        XSpreadsheet sheet = getSpreadsheetByIndex(doc, 0);
        XCell cell = null;
        try {
            cell = sheet.getCellByPosition(1, 1);
            XPropertySet cellProps =
                    (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, cell);
            Size size =
                    (Size) UnoRuntime.queryInterface(
                    Size.class, cellProps.getPropertyValue("Size"));
            double inchesWide = (double) size.Width / 2540d;
            double inchesHigh = (double) size.Height / 2540d;
            double DPI = 96;
            int dWide = (int) (inchesWide * DPI) + 1;
            int dHigh = (int) (inchesHigh * DPI) + 1;
            cellWidth = dWide;
            cellHeight = dHigh;
        } catch (UnknownPropertyException ex) {
            System.out.println(ex);
        } catch (WrappedTargetException ex) {
            System.out.println(ex);
        } catch (IndexOutOfBoundsException ex) {
            System.out.println(ex);
        }
    }

     public boolean toggleListen() {
        if (isListen() == true) {
            setListen(false);
        } else {
            setListen(true);
        }
        return isListen();
    }

    public void setListen(boolean listen) {
        this.listen = listen;
        XModel model = m_xFrame.getController().getModel();
        XController controller = model.getCurrentController();
        XSelectionSupplier selectionSupplier =
                (XSelectionSupplier) UnoRuntime.queryInterface(
                XSelectionSupplier.class, controller);
        if (listen) {
            selectionSupplier.addSelectionChangeListener(eventListener);
        } else {
            selectionSupplier.removeSelectionChangeListener(eventListener);
        }
    }

    public boolean isListen() {
        return listen;
    }

//    public boolean toggleConnect() {
//        if (isConnect() == true) {
//            setConnect(false);
//        } else {
//            setConnect(true);
//        }
//        return isConnect();
//    }

//    public void setConnect(boolean connect) {
//        this.connect = connect;
//    }

    public boolean isConnect() {
        return CONNECT;//connect;
    }

    // get from background color palette
    public Integer getCurrentColorInteger() {
        return Integer.valueOf(11164057);
    }
    // called by dispatch to reset dirty flag
    private void resetFirstTouch() {
        dirtyFlag = false;
    }
    
////////////////////////////////////////////////////////////////////////////////
///   End user mods, rest is netbeans OpenOffice Plugin AddOn boilerplate    ///
///   Note below addStatusListener and removeStatusListener not implmented   ///
///   Note below them, queryDispath aURL.Path comparison vs "FingerPaint"    ///
////////////////////////////////////////////////////////////////////////////////
    public void addStatusListener(com.sun.star.frame.XStatusListener xCtrl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

    public void removeStatusListener(com.sun.star.frame.XStatusListener xCtrl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

      // com.sun.star.frame.XDispatchProvider:
    public XDispatch queryDispatch(com.sun.star.util.URL aURL,
            String sTargetFrameName,
            int iSearchFlags) {
        if (aURL.Protocol.compareTo("org.openoffice.myaddon:") == 0) {
            if (aURL.Path.compareTo("FingerPaint") == 0) {
                return this;
            }
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch[] queryDispatches(
            com.sun.star.frame.DispatchDescriptor[] seqDescriptors) {
        int nCount = seqDescriptors.length;
        com.sun.star.frame.XDispatch[] seqDispatcher =
                new com.sun.star.frame.XDispatch[seqDescriptors.length];
        for (int i = 0; i < nCount; ++i) {
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,
                    seqDescriptors[i].FrameName,
                    seqDescriptors[i].SearchFlags);
        }
        return seqDispatcher;
    }

    public static XSingleComponentFactory __getComponentFactory(
            String sImplementationName) {
        XSingleComponentFactory xFactory = null;
        if (sImplementationName.equals(m_implementationName)) {
            xFactory = Factory.createComponentFactory(
                    myAddOn.class, m_serviceNames);
        }
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo(XRegistryKey xRegistryKey){
        return Factory.writeRegistryServiceInfo(m_implementationName,
                m_serviceNames,
                xRegistryKey);
    }

    // com.sun.star.lang.XInitialization:
    public void initialize(Object[] object)
            throws com.sun.star.uno.Exception {
        if (object.length > 0) {
            m_xFrame = (com.sun.star.frame.XFrame) UnoRuntime.queryInterface(
                    com.sun.star.frame.XFrame.class, object[0]);
        }
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
        return m_implementationName;
    }

    public boolean supportsService(String sService) {
        int len = m_serviceNames.length;
        for (int i = 0; i < len; i++) {
            if (sService.equals(m_serviceNames[i])) {
                return true;
            }
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }


}

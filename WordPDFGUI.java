import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.dnd.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.List;


import com.spire.doc.Document;
import com.spire.doc.ToPdfParameterList;
import com.spire.pdf.PdfConformanceLevel;


public class WordPDFGUI extends JFrame {
    private JButton convertButton;
    private JPanel mainPanel;
    private JTextField wordFilePath;
    private JButton chooseFileButton;
    private JTextField pdfFilePath;
    private JButton outputPathButton;
    private static File targetFile;
    private static String outputPath = System.getProperty("user.home") + "/Desktop";

    public WordPDFGUI(String title) {
        super(title);

        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setContentPane(mainPanel);


        MyDragDropListener myDragDropListener = new MyDragDropListener();
        new DropTarget(wordFilePath, myDragDropListener);

        wordFilePath.setEditable(false);
        pdfFilePath.setEditable(false);

        pdfFilePath.setText(outputPath);

        this.pack();

        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    convertToPdf(targetFile); // converted pdf is saved to desktop by default
                }
                catch (Exception ex) {
                    JOptionPane.showMessageDialog(mainPanel,"File not found");
                    /* wrote "Exception" instead of "NullPointerException" because
                    convertToPdf throws NullPointerException for trying to call
                    a method on a null targetFile instance
                    */
                }
            }
        });



        chooseFileButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter(".docx", "docx"));
                fc.setAcceptAllFileFilterUsed(false);
                int returnVal = fc.showDialog(null, "Open");



                if (returnVal == JFileChooser.APPROVE_OPTION && fc.getSelectedFile().getPath().endsWith(".docx")) {
                    targetFile = fc.getSelectedFile();
                    wordFilePath.setText(fc.getSelectedFile().getPath());
                } else if (!fc.getSelectedFile().getPath().endsWith(".docx")) {
                    JOptionPane.showMessageDialog(null, "Please choose a valid file");
                }
            }
        });

        outputPathButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser dc = new JFileChooser(); // dc = directory chooser
                dc.setCurrentDirectory(new java.io.File(outputPath));
                dc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                dc.setAcceptAllFileFilterUsed(false);

                if (dc.showDialog(null, "Open") == JFileChooser.APPROVE_OPTION) {
                    outputPath = dc.getCurrentDirectory().getPath();
                    pdfFilePath.setText(outputPath);
                }
            }
        });
    }

    public static void main(String[] args) {
        JFrame frame = new WordPDFGUI("WordToPDF");
        frame.setSize(650,450);
        frame.setVisible(true);

    }

    public class MyDragDropListener implements DropTargetListener {

        @Override
        public void drop(DropTargetDropEvent event) {

            // Accept copy drops
            event.acceptDrop(DnDConstants.ACTION_COPY);

            // Get the transfer which can provide the dropped item data
            Transferable transferable = event.getTransferable();

            // Get the data formats of the dropped item
            DataFlavor[] flavors = transferable.getTransferDataFlavors();

            // Loop through the flavors
            for (DataFlavor flavor : flavors) {

                try {

                    // If the drop items are files
                    if (flavor.isFlavorJavaFileListType()) {

                        // Get all of the dropped files
                        List<File> files = (List) transferable.getTransferData(flavor);

                        // Loop them through
                        for (File file : files) {

                            targetFile = file;
                            wordFilePath.setText(file.getPath());

                        }

                    }

                } catch (Exception e) {

                    // Print out the error stack
                    e.printStackTrace();

                }
            }

            // Inform that the drop is complete
            event.dropComplete(true);

        }

        @Override
        public void dragEnter(DropTargetDragEvent event) {
        }

        @Override
        public void dragExit(DropTargetEvent event) {
        }

        @Override
        public void dragOver(DropTargetDragEvent event) {
        }

        @Override
        public void dropActionChanged(DropTargetDragEvent event) {
        }
    }

    public static void convertToPdf(File file) throws FileNotFoundException {

        FileInputStream fis = new FileInputStream(file.getPath());

        Document doc = new Document(fis);

        ToPdfParameterList parameterList = new ToPdfParameterList();

        parameterList.setPdfConformanceLevel(PdfConformanceLevel.Pdf_A_1_A);

        doc.saveToFile(outputPath + "/" + file.getName().substring(0,file.getName().indexOf(".")) + ".pdf", parameterList);

    }
}


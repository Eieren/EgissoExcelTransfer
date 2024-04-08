package eieren.exceltoexcel;

import com.formdev.flatlaf.FlatIntelliJLaf;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

public class ExcelToExcel {

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(new FlatIntelliJLaf());
        } catch (UnsupportedLookAndFeelException e) {
            try {
                UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
            } catch (ClassNotFoundException | IllegalAccessException | InstantiationException | UnsupportedLookAndFeelException ex) {
            }
        }

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                MainScreen.main(args);
            }
        });
    }
}

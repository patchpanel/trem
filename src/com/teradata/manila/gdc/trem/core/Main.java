package com.teradata.manila.gdc.trem.core;

import com.teradata.manila.gdc.trem.gui.Logon;
import de.javasoft.plaf.synthetica.SyntheticaPlainLookAndFeel;

import javax.swing.*;
import java.awt.*;
import java.text.ParseException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class Main {

    // --Commented out by Inspection (5/2/2016 8:34 PM):private static final Logger LOG = Logger.getLogger(Main.class.getName());

    /**
     *
     * @param args
     */
    public static void main(String args[]) {

        boolean mode = false;

        //String userName2 = System.getProperty("user.name");
        //System.out.println(userName);
        //System.out.println(userName2);
        //System.exit(0);
        if (args.length > 0) {
            if ("admin".equals(args[0])) {
                System.out.println("Entering Administrative mode");
                mode = true;
            } else {
                System.out.println("Entering Normal mode");
            }
        }
        try {
            javax.swing.UIManager.setLookAndFeel(new SyntheticaPlainLookAndFeel());
        } catch (UnsupportedLookAndFeelException | ParseException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }

        /* Create and display the dialog */
        final boolean opMode = mode;
        java.awt.EventQueue.invokeLater(() -> {
            Logon dialog = new Logon(new javax.swing.JFrame(), true, opMode);
            dialog.addWindowListener(new java.awt.event.WindowAdapter() {
                @Override
                public void windowClosing(java.awt.event.WindowEvent e) {
                    System.exit(0);
                }
            });
            dialog.setSize(new Dimension(293,179));
            dialog.setLocationRelativeTo(null);
            dialog.setVisible(true);
        });
    }
}

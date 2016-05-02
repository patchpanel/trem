/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.util;

import java.io.IOException;
import java.io.OutputStream;
import java.util.logging.Logger;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;

/**
 * This class extends from OutputStream to redirect output to a JTextArrea
 *
 * @author www.codejava.net
 *
 */
public class CustomOutputStream extends OutputStream {
    private static final Logger LOG = Logger.getLogger(CustomOutputStream.class.getName());

    private final JTextArea textArea;

    /**
     *
     * @param textArea
     */
    public CustomOutputStream(JTextArea textArea) {
        this.textArea = textArea;
    }

    @Override
    public void write(int b) throws IOException {
        // redirects data to the text area
        //textArea.append(String.valueOf((char) b));
        // scrolls the text area to the end of data
        SwingUtilities.invokeLater(() -> {
            // redirects data to the text area
            textArea.append(String.valueOf((char) b));
            textArea.setCaretPosition(textArea.getDocument().getLength());
        });
    }
};

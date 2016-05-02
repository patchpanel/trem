package com.teradata.manila.gdc.trem.security;

import java.util.Hashtable;
import java.util.logging.Logger;
import javax.naming.Context;
import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import javax.naming.directory.DirContext;
import javax.naming.directory.InitialDirContext;

/**
 *
 * @author jl186034
 */
public class Authenticator {

    private String _userName;
    private String _passWord;
    private Exception _ex;

    public Exception getEx() {
        return _ex;
    }

    private void setEx(Exception _ex) {
        this._ex= _ex;
    }

    /**
     *
     */
    public Authenticator() {
        super();
    }

    /**
     *
     * @param username
     * @param password
     */
    public Authenticator(String username, String password) {
        this._userName = username;
        this._passWord = password;
    }

    /**
     *
     * @param _userName
     */
    public void setUserName(String _userName) {
        this._userName = _userName;
    }

    /**
     *
     * @param _passWord
     */
    public void setPassWord(String _passWord) {
        this._passWord = _passWord;
    }

    /**
     *
     * @return
     */
    public boolean authenticate() {
        Hashtable<String, String> env = new Hashtable<>();

        env.put(Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory");
        env.put(Context.PROVIDER_URL, "ldap://td.teradata.com:389");
        env.put(Context.SECURITY_AUTHENTICATION, "simple");
        env.put(Context.SECURITY_PRINCIPAL, "cn=" + this._userName + ",OU=APJ,OU=User Accounts,DC=TD,DC=TERADATA,DC=COM");
        env.put(Context.SECURITY_CREDENTIALS, this._passWord);
        DirContext ctx = null;
        NamingEnumeration<?> results = null;
        try {
            ctx = new InitialDirContext(env);
            return true;
        } catch (NamingException e) {
            this.setEx(e);
            System.out.println(e);
            return false;
        } finally {
            if (results != null) {
                try {
                    results.close();
                } catch (Exception e) {
                    this.setEx(e);
                    System.out.println(e);
                    return false;
                }
            }
            if (ctx != null) {
                try {
                    ctx.close();
                } catch (Exception e) {
                    this.setEx(e);
                    System.out.println(e);
                    return false;
                }
            }
        }
    }
    private static final Logger LOG = Logger.getLogger(Authenticator.class.getName());
}

import java.util.Base64;
import java.util.UUID;
import java.security.SecureRandom;
/**
 * Created by austin.calkins on 5/31/2017.
 */
public class tokenGenerator {

    protected static SecureRandom random = new SecureRandom();

    //function  returns a string with the username, and the password token generated.
    public static synchronized String generateToken( String username ) {
        long longToken = Math.abs( random.nextLong() );
        String random = Long.toString( longToken, 16 );
        return ( username + ":" + random );
    }

    public static void main(String[]args){

        System.out.println(tokenGenerator.generateToken("austin.calkins"));
    }

}
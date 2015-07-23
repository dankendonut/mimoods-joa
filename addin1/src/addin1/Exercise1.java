package addin1;

import java.util.Scanner;
import javax.swing.*;

import javax.swing.JOptionPane;


public class Exercise1 {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Scanner input = new Scanner(System.in);
		int number1 = 0;
		double number2=0 , sum;
		String name;
		sum=number1+number2;
		
		System.out.print("Enter an integer: ");
		number1 = input.nextInt();
		System.out.print("\n"+"Enter a floating point number: ");
		number2= input.nextDouble();
		
		System.out.print("\n"+"Enter your name: ");
		name= input.next();
		
		System.out.print("\n"+"Hi! "+name+", the sum of "+number1+" and "+number2+" is "+sum);
		
	}

}
package Test;

import java.util.Arrays;

public class SortingArray {

	    public static void main(String[] args) {
	        // Insertion sort

	        int [] array = {3,8,3,5,3,3,1,-10,12,78,34,0};
	       
	        // output 1 3 5 6 8
	        
	       


	        for(int i=1; i<array.length; i++){

	            int j= i-1; //0
	            int point = array[i];//8

	            while(j>=0 && array[j]>point) {

	                array[j + 1] = array[j];
	                j = j - 1;
	            }

	            array[j+1] = point;
	        }
	        Arrays.sort(array);
	        for(int n=0; n<array.length; n++){
	            System.out.println(array[n]);
	      
	        }
	        
	    }





	}


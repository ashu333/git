package com.practice.SpringDemo;

import org.springframework.context.support.ClassPathXmlApplicationContext;

public class SpringMainApp {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		//load the configuration file
		
		ClassPathXmlApplicationContext context = new ClassPathXmlApplicationContext("applicationContext.xml");
		
		//calling beans from configuration file
		// Spring inversion of control (IOC)
		Coach theCoach = context.getBean("coach", Coach.class);
		//calling methods of those beans loaded 
		
		System.out.println(theCoach.getDailyWorkout());
		
		//close the context..
		
		System.out.println(theCoach.getDailyFortune());
		
		context.close();
	}

}

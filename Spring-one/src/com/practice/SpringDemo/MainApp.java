package com.practice.SpringDemo;

public class MainApp {

	public static void main(String[] args) {
		
		Coach coach = new HockeyCoach();
		System.out.println(coach.getDailyWorkout());
	}
}

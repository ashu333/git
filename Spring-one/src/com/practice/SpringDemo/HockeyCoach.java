package com.practice.SpringDemo;

public class HockeyCoach implements Coach {
	
	FortuneService theFortuneService;
	
	public HockeyCoach() {
		
	}
	public HockeyCoach(FortuneService theFortuneService) {
		this.theFortuneService = theFortuneService;
	}
	
	@Override
	public String getDailyWorkout() {
		// TODO Auto-generated method stub
		return "run the groud in circular manner for hockey...";
	}
	@Override
	public String getDailyFortune() {
		return theFortuneService.getDailyFortune() + "this is for hockey class implementation.";
	}
	
	
}

package com.example.transmission.repository;

import com.example.transmission.domain.MonthlyIncidentSummary;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface MonthlyIncidentSummaryRepository extends JpaRepository<MonthlyIncidentSummary, Long> {

}

package com.example.transmission.repository;

import com.example.transmission.domain.ServiceSubscriber;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

@Repository
public interface ServiceSubscriberRepository extends JpaRepository<ServiceSubscriber, Long> {

    @Query(value =
            "SELECT " +
                    "complaint_group_progress_kpi_3h AS progress_3h, " +
                    "complaint_group_progress_kpi_24h AS progress_24h, " +
                    "complaint_group_progress_kpi_48h AS progress_48h " +
                    "FROM rp_service_subscribers " +
                    "LIMIT 1",
            nativeQuery = true)
    List<Map<String, Object>> getProgressKpi();

}

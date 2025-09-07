package com.example.transmission.repository;

import com.example.transmission.domain.ProvinceSubscriber;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

@Repository
public interface ProvinceSubscriberRepository extends JpaRepository<ProvinceSubscriber, Long> {

    @Query(value = "SELECT province_name, total_subscribers " +
            "FROM rp_province_subscribers " +
            "UNION ALL " +
            "SELECT 'TỔNG' AS province_name, SUM(total_subscribers) AS total_subscribers " +
            "FROM rp_province_subscribers", nativeQuery = true)
    List<Map<String, Object>> getProvinceSubscribers();
}
